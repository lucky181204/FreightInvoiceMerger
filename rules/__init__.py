"""Rules package — plugin-style rule discovery.

All rule modules in this package are auto-discovered by RuleRegistry.
Add new rules by creating rule_v2.py, rule_v3.py etc. — no code changes needed elsewhere.
"""

from rules.registry import RuleRegistry

__all__ = ["RuleRegistry"]
