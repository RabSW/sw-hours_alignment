FROM python:3.11-slim

# Install ODBC Driver for SQL Server
RUN apt-get update && apt-get install -y \
    curl \
    apt-transport-https \
    gnupg2 \
    unixodbc \
    unixodbc-dev \
    && curl https://packages.microsoft.com/keys/microsoft.asc | apt-key add - \
    && curl https://packages.microsoft.com/config/debian/11/prod.list > /etc/apt/sources.list.d/mssql-release.list \
    && apt-get update \
    && ACCEPT_EULA=Y apt-get install -y msodbcsql17 \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy project files
COPY pyproject.toml uv.lock ./
COPY src/ ./src/

# Install uv and dependencies
RUN pip install --no-cache-dir uv && uv sync

# Run the application
CMD ["uv", "run", "python", "src/main.py"]

