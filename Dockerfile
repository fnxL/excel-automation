# 
FROM python:3.12-alpine

# Install build tools and dependencies
RUN apk add --no-cache gcc musl-dev build-base

# 
WORKDIR /code

# 
COPY ./requirements.txt /code/requirements.txt

# 
RUN pip install --no-cache-dir --upgrade -r /code/requirements.txt

RUN pip install --upgrade pymupdf

# 
COPY ./ /code/

# 
CMD ["fastapi", "run", "app/main.py", "--port", "8000", "--proxy-headers"]