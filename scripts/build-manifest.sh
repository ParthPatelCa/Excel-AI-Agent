#!/bin/bash

echo "🔧 Building Office Add-in manifest..."

# Check if deployment URL is set
if [ -z "$DEPLOYMENT_URL" ]; then
    echo "❌ DEPLOYMENT_URL environment variable not set"
    echo "Please set it to your deployment URL (e.g., https://your-app.vercel.app)"
    exit 1
fi

# Replace placeholder URLs in manifest
sed "s|https://your-deployment-url.com|$DEPLOYMENT_URL|g" manifest.xml.template > manifest.xml

echo "✅ Manifest updated with deployment URL: $DEPLOYMENT_URL"
echo "📄 You can now sideload manifest.xml into Excel"
