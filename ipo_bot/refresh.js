const refreshButton = document.getElementById('refresh');
const refreshStatus = document.getElementById('refresh-status');
const api = 'https://api.github.com/repos/Tushar-Mishra999/ipo-bot';

refreshButton.addEventListener('click', async () => {
  let token = sessionStorage.getItem('ipo_refresh_token');
  if (!token) {
    token = window.prompt('Enter a GitHub fine-grained token for this repository with Actions read/write permission. It will be kept only in this browser tab.');
    if (!token) return;
    token = token.trim();
    sessionStorage.setItem('ipo_refresh_token', token);
  }
  refreshButton.disabled = true;
  refreshStatus.textContent = 'Starting refresh…';
  const headers = {Authorization: `Bearer ${token}`, Accept: 'application/vnd.github+json'};
  try {
    const started = new Date().toISOString();
    const response = await fetch(`${api}/actions/workflows/ipo-bot.yml/dispatches`, {
      method: 'POST', headers: {...headers, 'Content-Type': 'application/json'},
      body: JSON.stringify({ref: 'master'})
    });
    if (!response.ok) throw new Error(response.status === 401 || response.status === 403 ?
      'GitHub rejected the token. Check its Actions permission and try again.' :
      `Could not start refresh (${response.status}).`);
    refreshStatus.textContent = 'Refreshing IPO data…';
    for (let attempt = 0; attempt < 60; attempt++) {
      await new Promise(resolve => setTimeout(resolve, 10000));
      const runsResponse = await fetch(`${api}/actions/workflows/ipo-bot.yml/runs?event=workflow_dispatch&per_page=5`, {headers});
      if (!runsResponse.ok) throw new Error('Could not check refresh progress.');
      const runs = (await runsResponse.json()).workflow_runs;
      const run = runs.find(item => Date.parse(item.created_at) >= Date.parse(started) - 5000 && item.head_branch === 'master');
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
    if (error.message.includes('rejected the token')) sessionStorage.removeItem('ipo_refresh_token');
  } finally {
    refreshButton.disabled = false;
  }
});
