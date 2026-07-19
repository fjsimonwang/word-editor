# Zero-dependency Node editor + backend. No npm install needed.
FROM node:22-alpine

WORKDIR /app

COPY server.js ./
COPY public ./public

# Persist documents inside the container (mount a named volume to keep them).
RUN mkdir -p /app/data

ENV HOST=0.0.0.0
ENV PORT=3001
ENV DATA_DIR=/app/data

EXPOSE 3001

CMD ["node", "server.js"]
