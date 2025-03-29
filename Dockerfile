# 
FROM python:3.12-alpine

# 
WORKDIR /code

# 
COPY ./requirements.txt /code/requirements.txt

# 
RUN pip install --no-cache-dir --upgrade -r /code/requirements.txt

RUN pip install --upgdrade pymupdf

# 
COPY ./ /code/

# 
CMD ["fastapi", "run", "app/main.py", "--port", "8000", "--proxy-headers"]