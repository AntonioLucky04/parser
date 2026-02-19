FROM python:3.9-slim

# 1. Устанавливаем системные зависимости
RUN apt-get update && apt-get install -y \
    wget \
    unzip \
    libnss3 \
    libxss1 \
    libasound2 \
    libatk-bridge2.0-0 \
    libgtk-3-0 \
    libgbm-dev \
    python3-dev \
    gcc \
    g++ \
    curl \
    && rm -rf /var/lib/apt/lists/*

# 2. Устанавливаем Chrome (последняя стабильная версия)
RUN wget -q -O /tmp/chrome.deb https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb \
    && apt-get update \
    && apt-get install -y /tmp/chrome.deb \
    && rm /tmp/chrome.deb

# 3. Получаем версию Chrome
RUN echo "Installed Chrome:" && google-chrome --version

# 4. Скачиваем ChromeDriver из нового хранилища Chrome for Testing
# Сначала пробуем получить последнюю стабильную версию
RUN CHROME_VERSION=$(google-chrome --version | awk '{print $3}') \
    && echo "Chrome version: $CHROME_VERSION" \
    && echo "Downloading ChromeDriver for Chrome $CHROME_VERSION" \
    && wget -q -O /tmp/chromedriver.zip "https://storage.googleapis.com/chrome-for-testing-public/$CHROME_VERSION/linux64/chromedriver-linux64.zip" \
    && unzip /tmp/chromedriver.zip -d /tmp/ \
    && mv /tmp/chromedriver-linux64/chromedriver /usr/local/bin/chromedriver \
    && chmod +x /usr/local/bin/chromedriver \
    && rm -rf /tmp/chromedriver* \
    || (echo "Failed to download exact version, trying latest..." \
        && wget -q -O /tmp/chromedriver-latest.zip "https://storage.googleapis.com/chrome-for-testing-public/latest/linux64/chromedriver-linux64.zip" \
        && unzip /tmp/chromedriver-latest.zip -d /tmp/ \
        && mv /tmp/chromedriver-linux64/chromedriver /usr/local/bin/chromedriver \
        && chmod +x /usr/local/bin/chromedriver \
        && rm -rf /tmp/chromedriver*)

# 5. Проверяем ChromeDriver
RUN echo "ChromeDriver version:" && chromedriver --version

COPY ./stat/config.toml /app/stat/config.toml

WORKDIR /app
COPY Парсерсулучшеннымконфигом.py .

# 6. Устанавливаем Python библиотеки
RUN pip install --no-cache-dir \
    aiogram==3.10.0 \
    selenium==4.20.0 \
    beautifulsoup4==4.12.3 \
    pandas==2.2.2 \
    openpyxl==3.1.2 \
    python-docx==1.1.0 \
    PyPDF2==3.0.1 \
    lxml==5.2.1 \
    toml==0.10.2 \
    python-telegram-bot==20.7

CMD ["python", "Парсерсулучшеннымконфигом.py"]
