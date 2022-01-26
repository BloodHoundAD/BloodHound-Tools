FROM python:3.10-slim as build-stage

WORKDIR /tmp/code

COPY . .

RUN pip wheel --wheel-dir ./dist .

FROM python:3.10-slim

WORKDIR /app

COPY --from=build-stage /tmp/code/dist .

RUN pip install --no-cache-dir --no-index --find-links . bloodhound-tools