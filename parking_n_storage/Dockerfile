#specify the base image
FROM python:3.13.0

#tell the builder to copy files from the host and put them into the container image
COPY . ./parking_and_storage

#specify the path in the image where files will be copied and commands will be executed
WORKDIR /parking_and_storage

#tell the builder to run the specified command
RUN pip install -r requirements.txt \
    && mkdir ./files

#set the default command a container using this image will run
CMD ["python", "main.py"]