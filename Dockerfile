# ---- Stage 1: Build React frontend ----
FROM node:20-alpine AS frontend-build
WORKDIR /app/client
COPY Smart-BPML/client/package*.json ./
RUN npm ci
COPY Smart-BPML/client/ ./
RUN npm run build

# ---- Stage 2: Production ----
FROM node:20-slim
WORKDIR /app

# Install Python 3 + openpyxl (required by add_chart.py)
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 python3-pip \
    && pip3 install --break-system-packages openpyxl \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Install server dependencies (production only)
COPY Smart-BPML/server/package*.json ./
RUN npm ci --omit=dev

# Copy server source + static knowledge base
COPY Smart-BPML/server/ ./

# Copy built frontend into server's public directory
COPY --from=frontend-build /app/client/dist ./public

ENV NODE_ENV=production
EXPOSE 4000

CMD ["node", "server.js"]
