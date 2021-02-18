FROM python:3-onbuild
COPY . /usr/src/app
WORKDIR /usr/src/app

COPY requirements.txt ./
RUN pip install -r requirements.txt
COPY opa_email.py ./
CMD ["python3", "./opa_email.py"]