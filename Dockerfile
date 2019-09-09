FROM ubuntu:latest
RUN apt-get update
RUN apt-get install -y python3 python3-dev python3-pip
RUN apt-get install -y coinor-cbc
RUN apt-get install -y -qq glpk-utils
COPY . /app
WORKDIR /app
RUN pip3 install -r requirements.txt
# test locally
CMD ["gunicorn", "-b","0.0.0.0:8080", "main:server"]
# deploy on GCP
# CMD exec gunicorn -b :$PORT main:server
