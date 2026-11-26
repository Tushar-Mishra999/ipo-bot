const refreshButton = document.getElementById('refresh');
const refreshStatus = document.getElementById('refresh-status');
const api = 'https://api.github.com/repos/Tushar-Mishra999/ipo-bot';
const refreshApi = window.IPO_AI_ENDPOINT || '';

refreshButton.addEventListener('click', async () => {
  refreshButton.disabled = true;
  refreshStatus.textContent = 'Starting refresh…';
  try {
    if (!refreshApi) throw new Error('Manual refresh is unavailable.');
    const response = await fetch(`${refreshApi}/refresh`, {method: 'POST'});
    const body = await response.json();
    if (!response.ok) throw new Error(body.error || 'Could not start refresh.');
    refreshStatus.textContent = 'Refreshing IPO data…';
    for (let attempt = 0; attempt < 60; attempt++) {
      await new Promise(resolve => setTimeout(resolve, 10000));
      const runsResponse = await fetch(`${api}/actions/workflows/ipo-bot.yml/runs?event=workflow_dispatch&per_page=5`);
      if (!runsResponse.ok) throw new Error('Could not check refresh progress.');
      const runs = (await runsResponse.json()).workflow_runs;
      const run = runs.find(item => Date.parse(item.created_at) >= Date.parse(body.started) - 5000 && item.head_branch === 'master');
      if (!run) continue;
      if (run.status !== 'completed') continue;
      if (run.conclusion !== 'success') throw new Error('Refresh failed. Check the repository workflow logs.');
      refreshStatus.textContent = 'Refresh complete. Loading updated dashboard…';
      window.location.reload();
      return;
    }
    throw new Error('Refresh is taking longer than expected. Try reloading the page shortly.');
  } catch (error) {
    refreshStatus.textContent = error.message;
  } finally {
    refreshButton.disabled = false;
  }
});
