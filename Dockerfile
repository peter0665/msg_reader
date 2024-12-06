FROM python:3.10-slim
LABEL maintainer="Parker Lee<Parker_Lee@wistron.com>"

RUN mkdir -p /app
WORKDIR /app

COPY ./requirements.txt /app
RUN pip install -r requirements.txt

COPY ./ /app

ENTRYPOINT python interface.py