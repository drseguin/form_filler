# Form Filler Docker Setup

This document explains how to run the Form Filler application using Docker.

## Prerequisites

- Docker installed on your system
- Docker Compose installed on your system

## Running the Application

1. Clone this repository:

```bash
git clone <repository-url>
cd form_filler
```

2. Build and start the Docker container:

```bash
docker-compose up -d
```

3. Access the application in your browser at: http://localhost:8501

## Data Persistence

The Docker setup mounts the following directories as volumes to ensure data persistence:

- `./excel`: For Excel files
- `./templates`: For template files
- `./json`: For JSON files
- `./ai`: For AI-related files
- `./logs`: For application logs

Any files placed in these directories from your host system will be available to the application inside the container.

## Configuration

The application's configuration file `config.json` is also mounted as a volume, allowing you to modify the configuration without rebuilding the container.

## Stopping the Application

To stop the Docker container:

```bash
docker-compose down
```

## Troubleshooting

### Viewing Logs

To view the logs of the running container:

```bash
docker-compose logs -f
```

### Container Shell Access

To access a shell inside the running container:

```bash
docker exec -it form_filler bash
```

## Rebuilding After Code Changes

If you make changes to the code, rebuild the Docker image:

```bash
docker-compose up -d --build
``` 