"""Explicit offline cleanup; only retired task intermediates are removed."""
from pathlib import Path
import argparse
import json
import sys
sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / 'src'))


def main():
    from wps_pipeline.maintenance import cleanup_tasks
    p = argparse.ArgumentParser(description='保留各阶段总稿，清理已完成三表任务的中间数据。默认只预览。')
    p.add_argument('root', nargs='?', type=Path, default=ROOT / 'outputs', help='任务目录或装有多轮任务的父目录；默认包内outputs')
    p.add_argument('--apply', action='store_true', help='执行清理；默认仅留阶段总稿，旧02/03不再接续')
    p.add_argument('--keep-latest', action='store_true', help='额外完整保留最新任务及其修订依赖链，便于继续反馈')
    args = p.parse_args()
    try:
        if not args.root.exists() and args.root == ROOT / 'outputs':
            result = dict(status='NOTHING_TO_CLEAN', root=str(args.root), cleaned_tasks=0,
                          note='包内尚无outputs目录；可传入实际任务的父目录。')
        else: result = cleanup_tasks(args.root, args.apply, args.keep_latest)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 2 if result['status'] == 'PARTIAL' else 0
    except Exception as exc:
        print(json.dumps(dict(status='ERROR', error=str(exc), exception=type(exc).__name__), ensure_ascii=False, indent=2))
        return 2


if __name__ == '__main__':
    raise SystemExit(main())
