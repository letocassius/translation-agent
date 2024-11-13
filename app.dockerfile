# Use the official Python base image from the Docker Hub
FROM python:3.9-slim

# Ensure that the binary files are not cached
ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1
ENV PROTOCOL_BUFFERS_PYTHON_IMPLEMENTATION=python

# Install any necessary dependencies
RUN apt-get update && \
    apt-get install -y build-essential && \
    apt-get install -y libssl-dev libffi-dev libxml2-dev libxslt1-dev zlib1g-dev && \
    apt-get clean

# Install Python dependencies
COPY requirements.txt .
RUN pip install --upgrade pip
RUN pip install -r requirements.txt

# Create a working directory
WORKDIR /usr/src/app

# Copy the application files into the Docker image
COPY . .

# Add the Streamlit configuration file
COPY .streamlit /usr/src/app/.streamlit

# Expose the port Streamlit will run on
EXPOSE 8501

# Run the Streamlit app
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
