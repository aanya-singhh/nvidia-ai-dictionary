# Creating the GitHub Page

## 1. Create the repository

1. Sign in to GitHub.
2. Select **New repository**.
3. Choose a repository name such as `nvidia-ai-dictionary`.
4. Make it **Public** if the page should be publicly accessible.
5. Create the repository with a `README.md`.

The expected repository URL is:

```text
https://github.com/YOUR-GITHUB-USER/nvidia-ai-dictionary.git
```

## 2. Publish the starting page

Upload `site/index.html` through the GitHub website, or use Git:

```powershell
git clone https://github.com/YOUR-GITHUB-USER/nvidia-ai-dictionary.git
Copy-Item .\site\index.html .\nvidia-ai-dictionary\index.html -Force
git -C .\nvidia-ai-dictionary add index.html
git -C .\nvidia-ai-dictionary commit -m "Add AI dictionary site"
git -C .\nvidia-ai-dictionary push origin main
```

Git Credential Manager will open a browser sign-in when authentication is
needed. Do not place a personal access token in a script or configuration file.

## 3. Enable GitHub Pages

1. Open the repository on GitHub.
2. Select **Settings**.
3. Select **Pages** under **Code and automation**.
4. Under **Build and deployment**, select **Deploy from a branch**.
5. Choose the `main` branch and `/ (root)` folder.
6. Select **Save**.

The public URL normally becomes:

```text
https://YOUR-GITHUB-USER.github.io/nvidia-ai-dictionary/
```

The first deployment may take several minutes.

## 4. Connect the weekly workflow

In `local_config.json`, set:

```json
{
  "publish_website": true,
  "github_repository_url": "https://github.com/YOUR-GITHUB-USER/nvidia-ai-dictionary.git",
  "github_branch": "main"
}
```

The weekly publisher:

1. Fetches and fast-forward pulls the configured branch.
2. Replaces only the repository's root `index.html`.
3. Creates no commit when the file is unchanged.
4. Commits with the newsletter end date.
5. Pushes the configured branch.

Publication stops on authentication errors, merge conflicts, or
non-fast-forward history. Resolve those conditions manually and rerun only
after reviewing the remote changes.

## 5. Verify every deployment

After production:

1. Confirm the new commit appears on the configured branch.
2. Open the public Pages URL.
3. Confirm the newest weekly date, section totals, and model count.
4. If the old page is cached, append a query such as `?v=YYYY-MM-DD`.
