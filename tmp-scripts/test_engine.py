import sys
sys.path.insert(0, '.')

from engine import Engine, Logger, VerbosityLevel

# Change level here to see more or less detail:
#   QUIET   — no console output during processing
#   NORMAL  — summary + all validation issues (default)
#   VERBOSE — + step-by-step: cells found, tables matched
#   DEBUG   — + every anchor attempted and why it was accepted or rejected
logger = Logger(level=VerbosityLevel.NORMAL)

engine = Engine()
result = engine.process(
    pattern_file='pattern-example-2.xlsx',
    data_file='excel-file-to-test_pattern-2.xlsx',
    logger=logger,
)

print('\n=== CELLS ===')
for k, v in result['cells'].items():
    print(f'  {k}: {repr(v)}')

print('\n=== TABLES ===')
for t in result['tables']:
    print(f"\n  [table {t['table_index']} / instance {t['instance_index']}]  anchor={t['anchor']}")
    for i, h in enumerate(t['headers']):
        print(f'    HEADER {i+1}: {h}')
    for i, d in enumerate(t['data']):
        print(f'    DATA {i+1}: {d}')
    for i, f in enumerate(t['footers']):
        print(f'    FOOTER {i+1}: {f}')

issues = logger.issues()
if issues:
    print(f'\n=== ISSUES ({len(issues)}) ===')
    for r in issues:
        loc = f'  {r.location}' if r.location else ''
        print(f'  [{r.severity}]{loc}  {r.message}')
        if r.field:
            print(f'    field: {r.field}')
        if r.hint:
            print(f'    hint:  {r.hint}')
