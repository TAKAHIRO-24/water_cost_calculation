FROM python:3.7-alpine3.9
RUN apk --no-cache add libstdc++ \
  && apk --no-cache --virtual pydeps add gcc \
    g++ \
    python3-dev \
    musl-dev \
    cython \
  && pip install openpyxl \
  && apk del --purge pydeps
CMD ["/bin/sh"]