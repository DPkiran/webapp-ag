const { app } = require('@azure/functions');

app.http('getConfig', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'config',
  handler: async () => {
    return {
      jsonBody: {
        pollIntervalSeconds: Number(process.env.POLL_INTERVAL_SECONDS || 20),
      },
    };
  },
});
