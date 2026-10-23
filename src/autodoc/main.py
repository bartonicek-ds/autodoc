import json
import logging
import sys

from autodoc.generate_docs import generate_docs
from autodoc.logger import log_to_file, setup_logging
from autodoc.parse_config import (
    Config,
    ConfigError,
    load_config,
    make_output_dir,
    parse_args,
)

logger = logging.getLogger(__name__)

LOG_FILE_NAME = "logs.txt"


def run(config: Config) -> None:
    logger.info("input dir:  %s", config.input_dir)
    logger.info("output dir: %s", config.output_dir)
    logger.info("vars:\n%s", json.dumps(config.vars, ensure_ascii=False, indent=2))
    generate_docs(config)


def main() -> None:
    setup_logging()
    try:
        input_dir = parse_args()
        # Start the log file before reading the input, so input problems end up in it
        output_dir = make_output_dir(input_dir)
        log_to_file(output_dir / LOG_FILE_NAME)

        config = load_config(input_dir)
        run(config)
    except ConfigError as e:
        logger.error("%s", e)
        sys.exit(1)
    except Exception:
        # Log crashes with their details, so they end up in the log file too
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
