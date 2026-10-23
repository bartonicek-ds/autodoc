import json
import logging
import sys

from autodoc.generate_docs import generate_docs
from autodoc.logger import log_to_file, setup_logging
from autodoc.parse_config import Config, parse_config

logger = logging.getLogger(__name__)

LOG_FILE_NAME = "logs.txt"


def run(config: Config) -> None:
    config.output_dir.mkdir(parents=True, exist_ok=True)
    log_to_file(config.output_dir / LOG_FILE_NAME)

    logger.info("input dir:  %s", config.input_dir)
    logger.info("output dir: %s", config.output_dir)
    logger.info("vars:\n%s", json.dumps(config.vars, ensure_ascii=False, indent=2))
    generate_docs(config)


def main() -> None:
    setup_logging()
    try:
        config = parse_config()
        run(config)
    except Exception:
        # Log crashes too, so their details end up in the log file
        logger.exception("unexpected error")
        sys.exit(1)
    finally:
        # A double-clicked executable gets its own console window, which on Windows
        # closes as soon as the program ends, so keep it open until the user is done
        is_frozen = getattr(sys, "frozen", False)
        if is_frozen and sys.stdin.isatty():
            input("\nStiskněte Enter pro zavření…")


if __name__ == "__main__":
    main()
