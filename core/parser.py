"""Invoice data parser — extracts values from a single invoice sheet."""

from rules.registry import RuleRegistry


def parse_invoice(sheet, rule_id: str) -> dict:
    """
    Parse a single invoice sheet (SheetWrapper) using the specified rule.
    Returns a dict mapping target columns to extracted values.
    """
    rule = RuleRegistry.get_rule(rule_id)
    if not rule:
        raise ValueError(f"Rule not found: {rule_id}")

    result = {}
    for mapping in rule.FIELD_MAPPINGS:
        target_col = mapping["target_col"]
        try:
            value = rule.extract_value(sheet, mapping)
            result[target_col] = value
        except Exception as e:
            result[target_col] = ""
            import logging
            logging.getLogger(__name__).warning(
                f"Failed to extract {mapping['source']} → col {target_col}: {e}"
            )

    return result
