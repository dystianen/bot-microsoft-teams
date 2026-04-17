module.exports = {
  apps: [
    {
      name: 'bot-microsoft-teams',
      script: 'src/bots/TelegramBot.js',
      instances: 1,
      autorestart: true,
      watch: false,
      max_memory_restart: '1G',
      kill_timeout: 300000, // Wait up to 5 minutes to finish current task
      env: {
        NODE_ENV: 'production',
        DISPLAY: ':22',
      },
    },
  ],
};
