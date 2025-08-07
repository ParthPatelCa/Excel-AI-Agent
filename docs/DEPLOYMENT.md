# 🚀 Deployment Guide

## Prerequisites

- Node.js 18+ installed
- OpenAI API key
- Git repository access
- Hosting platform account

## Quick Deploy Options

### 1. Vercel (Recommended for beginners)

[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new/clone?repository-url=https://github.com/ParthPatelCa/Excel-AI-Agent)

1. Click the deploy button above
2. Connect your GitHub account
3. Set environment variables:
   - `OPENAI_API_KEY`: Your OpenAI API key
   - `NODE_ENV`: production
4. Deploy!

### 2. Railway

1. Connect repository to Railway
2. Add environment variables
3. Deploy automatically

### 3. Render

1. Connect repository to Render
2. Choose "Web Service"
3. Set build command: `npm install`
4. Set start command: `npm start`
5. Add environment variables

### 4. Docker

```bash
# Build image
docker build -t excel-ai-agent .

# Run container
docker run -p 3000:3000 --env-file .env excel-ai-agent
```

## Post-Deployment Steps

1. **Update manifest.xml** with your deployment URL:
   ```bash
   export DEPLOYMENT_URL="https://your-app-url.com"
   ./scripts/build-manifest.sh
   ```

2. **Test the application** at your deployment URL

3. **Sideload the add-in** into Excel:
   - Excel Desktop: Insert → My Add-ins → Upload My Add-in
   - Excel Web: Insert → Office Add-ins → Upload My Add-in

4. **Test functionality** with sample data

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `OPENAI_API_KEY` | Yes | Your OpenAI API key |
| `NODE_ENV` | Yes | Set to 'production' |
| `PORT` | No | Server port (default: 3000) |
| `RATE_LIMIT_MAX_REQUESTS` | No | Rate limit (default: 100) |
| `MAX_DATA_ROWS` | No | Max data rows (default: 1000) |

## Troubleshooting

### Common Issues

1. **Add-in won't load**: Check that all URLs in manifest.xml use HTTPS
2. **CORS errors**: Verify your domain is in the allowed origins
3. **API errors**: Check OpenAI API key and rate limits

### Monitoring

Check these endpoints for health monitoring:
- `/health` - Health check
- `/api` - API documentation

## Security Considerations

- Never commit your OpenAI API key to version control
- Use environment variables for all secrets
- Enable rate limiting in production
- Monitor API usage and costs
