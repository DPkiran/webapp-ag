const { app } = require('@azure/functions');
const { fetchTableRows, clearCaches } = require('../graphClient');

app.http('getData', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'data',
  handler: async (request, context) => {
    try {
      const { headers, rows } = await fetchTableRows();
      return {
        jsonBody: { headers, rows, fetchedAt: new Date().toISOString() },
      };
    } catch (err) {
      context.error('Graph fetch failed:', err.response ? JSON.stringify(err.response.data) : err.message);
      clearCaches(); // self-heal if the site/file moved or IDs went stale
      return {
        status: 502,
        jsonBody: { error: 'Failed to read data from SharePoint' },
      };
    }
  },
});
