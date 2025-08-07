#!/bin/bash

set -e

echo "🚀 Starting deployment process..."

# Check if environment is set
if [ -z "$NODE_ENV" ]; then
    echo "❌ NODE_ENV not set"
    exit 1
fi

# Install dependencies
echo "📦 Installing dependencies..."
npm ci --only=production

# Build if needed
echo "🔨 Building application..."
npm run build

# Start application
echo "✅ Starting application..."
npm start

echo "🎉 Deployment completed successfully!"
