"""Generic epidemiological maths shared across skills.

Domain thresholds (e.g. what CFR counts as 'critical') stay in the skill;
only the raw, reusable calculations live here.
"""


def cfr_v(d, c):
    """Case fatality rate as a percentage, rounded to 1 dp. 0.0 if no cases."""
    return round(d / c * 100, 1) if c > 0 else 0.0


def pct_v(v, t):
    """v as a percentage of t, rounded to 1 dp. 0.0 if t == 0."""
    return round(v / t * 100, 1) if t > 0 else 0.0
