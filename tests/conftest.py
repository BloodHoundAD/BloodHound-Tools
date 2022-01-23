import logging

handler = logging.StreamHandler()
handler.setFormatter(
    logging.Formatter(
        style="{",
        fmt="[{name}:{filename}] {levelname} - {message}"
    )
)

log = logging.getLogger("bh-tools")
log.setLevel(logging.INFO)
log.addHandler(handler)
