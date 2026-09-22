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
      const detail = err.response ? err.response.data : err.message;
      context.error('Graph fetch failed:', JSON.stringify(detail));
      clearCaches(); // self-heal if the site/file moved or IDs went stale
      return {
        status: 502,
        jsonBody: { error: 'Failed to read data from SharePoint', detail },
      };
    }
  },
});




    
