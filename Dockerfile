FROM ubuntu:latest
RUN apt-get update
RUN apt-get install -y python3 python3-dev python3-pip
RUN apt-get install -y coinor-cbc
RUN apt-get install -y -qq glpk-utils
COPY . /app
WORKDIR /app
RUN pip3 install -r requirements.txt
ENTRYPOINT ["python3"]
CMD ["main.py"]
