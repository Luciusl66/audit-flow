"""Select reviewed account routes and read only their rule modules."""

from dataclasses import dataclass
from pathlib import Path, PurePosixPath
import re
from string import Template

from yaml import YAMLError, safe_load

from core.schema import AuditError, account_code, text_value

ROOT = Path(__file__).resolve().parents[1]
OPERATORS = {
    "debit_credit_balance": ("GL", "complete_vouchers", ("difference",)),
    "period_end_nonzero": ("TB", ("period_closed", "currency_exchange_closed"),
                           ("account_code", "balance", "debit_account", "credit_account", "amount")),
    "balance_direction": ("TB", None, ("account_code", "balance")),
    "reclassification": ("TB", None, ("account_code", "balance")),
    "contra_exceeds_gross": ("TB", None, ("balance", "gross", "excess")),
    "gl_keyword_alert": ("GL", None, ("account_code", "keyword")),
}
RULE_FIELDS = {
    "rule_id", "cas_ref", "severity", "target_scope", "account_pattern",
    "check_type", "condition", "suggested_entry", "message_template",
}


@dataclass(frozen=True)
class CasRoute:
    hard_rules: tuple[Path, ...]
    soft_refs: tuple[Path, ...]
    unmatched_codes: tuple[str, ...]
    named_codes: frozenset[str] = frozenset()


def routed_file(root, relative, directory, suffix):
    relative = PurePosixPath(relative)
    if relative.parent.as_posix() != directory or relative.suffix != suffix:
        raise ValueError
    path = root.joinpath(*relative.parts)
    if path.resolve() != path or not path.is_file():
        raise ValueError
    return path


def yaml_mapping(path):
    with path.open(encoding="utf-8") as file:
        document = safe_load(file)
    if not isinstance(document, dict) or type(document.get("schema_version")) is not int or document["schema_version"] != 1:
        raise ValueError
    return document


def confirmed_account_names(values):
    """Normalize codes supplied with explicitly verified chart names."""
    try:
        return {account_code(code): text_value(name) for code, name in (values or {}).items()}
    except (AttributeError, TypeError, ValueError):
        raise AuditError("CAS 科目名称映射无效。请核对账套提供的科目代码及名称。") from None


def account_hierarchy(code):
    boundaries = [match.start() for match in re.finditer(r"[._/-]", code)]
    return code, *(code[:index] for index in reversed(boundaries))


def route_cas(account_codes, *, account_names=None, project_root=ROOT):
    """Use a verified name, an exact account code, or a delimited ancestor."""
    root = project_root.resolve()
    confirmed_names = confirmed_account_names(account_names)
    try:
        config = yaml_mapping(routed_file(root, "rules/cas_router.yaml", "rules", ".yaml"))
        accounts = config["accounts"]
        named_accounts = config.get("named_accounts", {})
        global_route = config["global"]
        if (config["matching"] != "exact_then_delimited_parent"
                or not isinstance(accounts, dict) or not isinstance(named_accounts, dict)
                or not isinstance(global_route, dict)):
            raise ValueError
        selected = [global_route]
        unmatched = []
        named_codes = frozenset(code for code, name in confirmed_names.items() if name in named_accounts)
        for code in dict.fromkeys(account_code(value) for value in account_codes):
            candidates = account_hierarchy(code)
            named_code = next((candidate for candidate in candidates if candidate in named_codes), None)
            if named_code is not None:
                selected.append(named_accounts[confirmed_names[named_code]])
                continue
            key = next((candidate for candidate in candidates if candidate in accounts), None)
            if key is None:
                unmatched.append(code)
            else:
                selected.append(accounts[key])

        def files(field, directory, suffix):
            paths = []
            for route in selected:
                entries = route[field]
                if not isinstance(entries, list):
                    raise ValueError
                paths.extend(routed_file(root, entry, directory, suffix) for entry in entries)
            return tuple(dict.fromkeys(paths))

        return CasRoute(files("hard_rules", "rules/cas", ".yaml"),
                        files("soft_refs", "references/cas", ".md"), tuple(unmatched), named_codes)
    except (OSError, RuntimeError, YAMLError, KeyError, TypeError, ValueError):
        raise AuditError("CAS 路由配置无效。请核对科目编码和项目内规则或卡片路径。") from None


def matches_account(pattern, code, account_names=None):
    kind = pattern["type"]
    if kind == "all":
        return True
    if kind == "prefix":
        return code.startswith(pattern["value"])
    if kind == "names":
        return (account_names or {}).get(code) in pattern["values"]
    return re.fullmatch(pattern["value"], code) is not None


def verify_pattern(pattern):
    kind = pattern["type"]
    if kind == "all":
        if set(pattern) != {"type"}:
            raise ValueError
    elif kind == "names":
        names = pattern["values"]
        if (set(pattern) != {"type", "values"} or not isinstance(names, list) or not names
                or not all(text_value(name) == name for name in names) or len(set(names)) != len(names)):
            raise ValueError
    elif set(pattern) != {"type", "value"} or not pattern["value"]:
        raise ValueError
    elif kind == "regex":
        re.compile(pattern["value"])
    elif kind != "prefix" or account_code(pattern["value"]) != pattern["value"]:
        raise ValueError


def verify_condition(check, condition, pattern, context):
    if check == "gl_keyword_alert":
        keywords = condition["keywords"]
        if (set(condition) != {"field", "match", "keywords"} or condition["field"] != "summary"
                or condition["match"] != "contains" or not isinstance(keywords, list) or not keywords
                or not all(isinstance(word, str) and word.strip() for word in keywords)
                or len(set(keywords)) != len(keywords)):
            raise ValueError
        return
    keys = {"tolerance"}
    if context is not None:
        keys.add("requires_context")
        valid_contexts = context if isinstance(context, tuple) else (context,)
        if condition["requires_context"] not in valid_contexts:
            raise ValueError
    if check == "debit_credit_balance":
        keys.add("group_by")
        if pattern["type"] != "all" or condition["group_by"] != ["date", "voucher_no"]:
            raise ValueError
    if check == "period_end_nonzero" and "root_only" in condition:
        keys.add("root_only")
        if type(condition["root_only"]) is not bool:
            raise ValueError
    if check == "balance_direction":
        keys.add("expected_direction")
        if condition["expected_direction"] not in {"debit", "credit", "zero"}:
            raise ValueError
    if check == "reclassification":
        keys.update(("trigger_direction", "require_auxiliary"))
        if condition["trigger_direction"] not in {"debit", "credit"} or type(condition["require_auxiliary"]) is not bool:
            raise ValueError
    if check == "contra_exceeds_gross":
        keys.update(("gross_account_pattern", "grouping"))
        verify_pattern(condition["gross_account_pattern"])
        if (condition["grouping"] not in {"confirmed_scope", "auxiliary"}
                or condition["gross_account_pattern"]["type"] == "names"):
            raise ValueError
    if set(condition) != keys or not re.fullmatch(r"[0-9]+\.[0-9]{2}", condition["tolerance"]):
        raise ValueError


def verify_rule(rule, cas_no):
    if set(rule) != RULE_FIELDS or not re.fullmatch(fr"CAS{cas_no:02d}_[A-Z][A-Z0-9_]*", rule["rule_id"]):
        raise ValueError
    if rule["severity"] not in {"ERROR", "WARNING", "INFO"} or not rule["cas_ref"].strip():
        raise ValueError
    check = rule["check_type"]
    scope, context, variables = OPERATORS[check]
    if rule["target_scope"] != scope:
        raise ValueError
    pattern = rule["account_pattern"]
    verify_pattern(pattern)
    if pattern["type"] == "names" and check != "balance_direction":
        raise ValueError
    verify_condition(check, rule["condition"], pattern, context)

    values = dict.fromkeys(variables, "1.00")
    if not Template(rule["message_template"]).substitute(values).strip():
        raise ValueError
    entry = rule["suggested_entry"]
    if entry is not None:
        if set(entry) != {"type", "debit", "credit", "amount"} or entry["type"] not in {"RJE", "AJE"}:
            raise ValueError
        for field in ("debit", "credit"):
            if not Template(entry[field]).substitute(values).strip():
                raise ValueError
        if entry["amount"] is not None and not re.fullmatch(
            r"[0-9]+\.[0-9]{2}", Template(entry["amount"]).substitute(values)
        ):
            raise ValueError


def load_routed_rules(route):
    """Skip stub modules and reject invalid or unimplemented executable rules."""
    definitions = {}
    try:
        for path in route.hard_rules:
            module = yaml_mapping(path)
            cas_no = module["cas_no"]
            rules = module["rules"]
            if (set(module) != {"schema_version", "cas_no", "name", "status", "rules"}
                    or type(cas_no) is not int or not 0 <= cas_no <= 42
                    or not module["name"].strip() or not isinstance(rules, list)):
                raise ValueError
            if module["status"] == "stub" and not rules:
                continue
            if module["status"] != "distilled":
                raise ValueError
            for rule in rules:
                verify_rule(rule, cas_no)
                rule_id = rule["rule_id"]
                if rule_id in definitions:
                    raise ValueError
                definitions[rule_id] = rule
    except (OSError, YAMLError, AttributeError, KeyError, TypeError, ValueError, re.error):
        raise AuditError("CAS 规则模块无效，或包含尚未接入执行的算子。") from None
    return definitions
