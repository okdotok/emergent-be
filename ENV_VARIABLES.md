# Environment Variables Required for Deployment

This document lists all environment variables that need to be configured in your deployment platform (e.g., Render).

## Required Environment Variables

### Database
- `MONGO_URL`: MongoDB connection string
  - Example: `mongodb+srv://username:password@cluster.mongodb.net/?appName=Cluster0`
  
- `DB_NAME`: Database name (optional, defaults to `theglobal_uren`)

### CORS Configuration
- `CORS_ORIGINS`: Comma-separated list of allowed origins for CORS
  - Example: `https://your-frontend-domain.vercel.app`
  - Multiple origins: `https://domain1.com,https://domain2.com`

### Frontend URL
- `FRONTEND_URL`: Frontend application URL (used in email links)
  - Example: `https://your-frontend-domain.vercel.app`

### Security
- `JWT_SECRET_KEY`: Secret key for JWT token signing (optional, but recommended in production)
  - Use a strong random string in production

### Email Configuration (Optional)
- `SMTP_HOST`: SMTP server host (defaults to `smtp.transip.email`)
- `SMTP_PORT`: SMTP server port (defaults to `465`)
- `SMTP_USERNAME`: SMTP username
- `SMTP_PASSWORD`: SMTP password
- `SMTP_FROM`: Email sender address (defaults to `info@theglobal-bedrijfsdiensten.nl`)
- `SMTP_SECURE`: Connection security type - `ssl` or `tls` (defaults to `ssl`)

### Server Configuration
- `PORT`: Server port (Render sets this automatically, but defaults to `10000`)

## Setting Environment Variables on Render

1. Go to your Render service dashboard
2. Navigate to "Environment" section
3. Add each variable with its value
4. Save and redeploy

**Note:** Never commit `.env` files with actual credentials to version control.

