module.exports = async ({ github, context, core, token }) => {
  if (!token) {
    core.setFailed('COPILOT_ASSIGN_TOKEN is required to assign the coding agent.');
    return;
  }

  const { owner, repo } = context.repo;
  const { data } = await github.rest.repos.get({ owner, repo });
  await github.request('POST /repos/{owner}/{repo}/issues/{issue_number}/assignees', {
    owner,
    repo,
    issue_number: context.payload.issue.number,
    assignees: ['copilot-swe-agent[bot]'],
    agent_assignment: {
      target_repo: `${owner}/${repo}`,
      base_branch: data.default_branch,
      custom_instructions: '',
      custom_agent: '',
      model: ''
    },
    headers: { 'X-GitHub-Api-Version': '2022-11-28' }
  });
  core.info(`Assigned coding agent to issue #${context.payload.issue.number}.`);
};
