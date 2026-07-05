"""Rule registry — plugin-style rule discovery and registration.

New rules can be added by creating a new module in rules/ and registering
it here. No UI or core code changes needed.
"""

import importlib
import pkgutil
import rules


class RuleRegistry:
    """Singleton registry for all invoice processing rules."""

    _rules = {}
    _loaded = False

    @classmethod
    def discover(cls):
        """Auto-discover all rule modules in the rules package."""
        if cls._loaded:
            return cls._rules

        for importer, mod_name, is_pkg in pkgutil.iter_modules(rules.__path__):
            if mod_name.startswith("_"):
                continue
            try:
                mod = importlib.import_module(f"rules.{mod_name}")
                if hasattr(mod, "RULE_META") and (hasattr(mod, "FIELD_MAPPINGS") or hasattr(mod, "write_sort_data_fn")):
                    rule_id = mod.RULE_META["id"]
                    cls._rules[rule_id] = mod
            except Exception:
                import logging
                logging.getLogger(__name__).exception(f"Failed to load rule: {mod_name}")

        cls._loaded = True
        return cls._rules

    @classmethod
    def get_rule(cls, rule_id: str):
        """Get a rule module by ID."""
        cls.discover()
        return cls._rules.get(rule_id)

    @classmethod
    def list_rules(cls) -> list[dict]:
        """Return list of rule metadata dicts for UI display."""
        cls.discover()
        return [
            {
                "id": mod.RULE_META["id"],
                "name": mod.RULE_META["name"],
                "description": mod.RULE_META.get("description", ""),
                "template_name": mod.RULE_META.get("template_name", ""),
                "start_row": mod.RULE_META.get("start_row", 5),
                "sort_column": mod.RULE_META.get("sort_column", "U"),
            }
            for mod in cls._rules.values()
        ]

    @classmethod
    def register_rule(cls, rule_id: str, module):
        """Manually register a rule module."""
        cls._rules[rule_id] = module
