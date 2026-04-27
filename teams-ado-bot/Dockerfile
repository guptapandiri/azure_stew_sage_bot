FROM node:22-alpine

WORKDIR /app

COPY package*.json ./
RUN npm ci --omit=dev

COPY index.js ./
COPY src/ ./src/

EXPOSE 3979

CMD ["node", "index.js"]
