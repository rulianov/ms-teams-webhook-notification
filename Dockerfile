FROM node:18-alpine

WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm install --omit=dev

# Copy application code
COPY server.js .

# Expose port
EXPOSE 8080

# Run as non-root user
USER node

# Start the application
CMD ["node", "server.js"]