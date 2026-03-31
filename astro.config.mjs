import { defineConfig } from "astro/config";

const isGitHubPagesBuild = process.env.GITHUB_ACTIONS === "true";
const repository = process.env.GITHUB_REPOSITORY?.split("/")[1];
const repositoryOwner = process.env.GITHUB_REPOSITORY_OWNER;

export default defineConfig({
  site:
    isGitHubPagesBuild && repositoryOwner ? `https://${repositoryOwner}.github.io` : undefined,
  base: isGitHubPagesBuild && repository ? `/${repository}` : "/",
  server: {
    host: true,
  },
});
