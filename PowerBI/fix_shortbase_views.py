path = r"D:\Work\Driver Survey\PowerBI\create_views.sql"

with open(path, "r", encoding="utf-8") as f:
    sql = f.read()

changes = []

# ── Step 1: Replace the formatting expression in all 6 views ──────────────────
# These views source from vw_ShortBase (4-space indent) and re-format yearweek.
# vw_ShortBase now returns yearweek/yearweek_sort already formatted — just pass through.

old_fmt = (
    "    CAST(yearweek/100 AS VARCHAR) + '-' + RIGHT('0' + CAST(yearweek%100 AS VARCHAR), 2) AS yearweek,\n"
    "    yearweek AS yearweek_sort,"
)
new_fmt = "    yearweek,\n    yearweek_sort,"

count = sql.count(old_fmt)
assert count == 6, f"Expected 6 occurrences of format expression, found {count}"
sql = sql.replace(old_fmt, new_fmt)
changes.append(f"Replaced yearweek format expression in {count} views")

# ── Step 2: Fix GROUP BY yearweek → GROUP BY yearweek, yearweek_sort ──────────
# vw_WeeklySatisfaction, vw_WeeklyNPS, vw_IncentiveByWeek
old = "GROUP BY yearweek\nHAVING COUNT(*) >= 100"
new = "GROUP BY yearweek, yearweek_sort\nHAVING COUNT(*) >= 100"
count = sql.count(old)
assert count == 3, f"Expected 3 occurrences of 'GROUP BY yearweek HAVING >= 100', found {count}"
sql = sql.replace(old, new)
changes.append(f"Fixed GROUP BY yearweek → yearweek, yearweek_sort ({count} views with HAVING >= 100)")

# vw_SatisfactionByCityWeek: GROUP BY city, yearweek  HAVING >= 10
old = "GROUP BY city, yearweek\nHAVING COUNT(*) >= 10"
new = "GROUP BY city, yearweek, yearweek_sort\nHAVING COUNT(*) >= 10"
assert old in sql, "vw_SatisfactionByCityWeek GROUP BY not found"
sql = sql.replace(old, new, 1)
changes.append("Fixed vw_SatisfactionByCityWeek GROUP BY")

# vw_RideShareByCityWeek: GROUP BY city, yearweek  (no HAVING)
old = "GROUP BY city, yearweek\n;"
new = "GROUP BY city, yearweek, yearweek_sort\n;"
assert old in sql, "vw_RideShareByCityWeek GROUP BY not found"
sql = sql.replace(old, new, 1)
changes.append("Fixed vw_RideShareByCityWeek GROUP BY")

# vw_NavigationByWeek: GROUP BY yearweek, snapp_last_trip_navigation
old = "GROUP BY yearweek, snapp_last_trip_navigation"
new = "GROUP BY yearweek, yearweek_sort, snapp_last_trip_navigation"
assert old in sql, "vw_NavigationByWeek GROUP BY not found"
sql = sql.replace(old, new, 1)
changes.append("Fixed vw_NavigationByWeek GROUP BY")

with open(path, "w", encoding="utf-8") as f:
    f.write(sql)

print(f"Applied {len(changes)} changes:")
for c in changes:
    print(f"  OK {c}")
