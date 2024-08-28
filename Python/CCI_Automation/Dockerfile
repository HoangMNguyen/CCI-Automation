# Use the slim Python runtime as a parent image
FROM python:3.11-slim

# Set the working directory in the container
WORKDIR /app

# Copy the current directory contents into the container at /app
# COPY Setup /app/Setup

# Install system dependencies
RUN apt-get update && apt-get install -y \
    libglib2.0-0 \
    libgl1 \
    libegl1 \
    libxkbcommon-x11-0 \
    libglib2.0-bin \
    libgl1-mesa-glx \
    libgl1-mesa-dri \
    libxcb-xinerama0 \
    build-essential \
    libpulse0 \
    libx11-xcb1 \
    libfontconfig1 \
    libdbus-1-3 \
    libxcb-icccm4 \
    libxcb-image0 \
    libxcb-keysyms1 \
    libxcb-randr0 \
    libxcb-render-util0 \
    libxcb-render0 \
    libxcb-shape0 \
    libxcb-sync1 \
    libxcb-xfixes0 \
    libxcb-xinerama0 \
    libxcb-xkb1 \
    libxcb1 \
    libxrender1 \
    libxi6 \
    libxcb-cursor0 \
    xauth \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Install any needed packages specified in requirements.txt
# RUN pip install --trusted-host pypi.python.org -r /app/Setup/requirements.txt

# ENV DISPLAY=172.26.48.1:0.0

# Copy the rest of the application code into the container at /app
COPY . /app

# Expose the port that the app runs on
# EXPOSE 5000

# Run check.py when the container launches
CMD ["python", "check.py"]