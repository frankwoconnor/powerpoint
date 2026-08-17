#!/usr/bin/env bash

set -euo pipefail

KEY_PATH="/Users/foconnor/.ssh/id_ed25519"
PUBLIC_KEY_PATH="${KEY_PATH}.pub"
SSH_CONFIG="/Users/foconnor/.ssh/config"
REMOTE_URL="git@github.com:frankwoconnor/powerpoint.git"

echo "Configuring GitHub SSH access..."

# Confirm that the generated key pair exists.
if [[ ! -f "$KEY_PATH" ]]; then
  echo "ERROR: Private key not found at:"
  echo "  $KEY_PATH"
  exit 1
fi

if [[ ! -f "$PUBLIC_KEY_PATH" ]]; then
  echo "ERROR: Public key not found at:"
  echo "  $PUBLIC_KEY_PATH"
  exit 1
fi

# Ensure the SSH directory and key files have secure permissions.
mkdir -p "/Users/foconnor/.ssh"
chmod 700 "/Users/foconnor/.ssh"
chmod 600 "$KEY_PATH"
chmod 644 "$PUBLIC_KEY_PATH"

# Create a backup of an existing SSH configuration.
if [[ -f "$SSH_CONFIG" ]]; then
  cp "$SSH_CONFIG" "${SSH_CONFIG}.backup.$(date +%Y%m%d-%H%M%S)"
fi

touch "$SSH_CONFIG"
chmod 600 "$SSH_CONFIG"

# Add the GitHub configuration only if it is not already present.
if ! grep -qE '^[[:space:]]*Host[[:space:]]+github\.com[[:space:]]*$' "$SSH_CONFIG"; then
  cat >> "$SSH_CONFIG" <<'EOF'

Host github.com
  HostName github.com
  User git
  AddKeysToAgent yes
  UseKeychain yes
  IdentityFile /Users/foconnor/.ssh/id_ed25519
  IdentitiesOnly yes
EOF

  echo "Added GitHub configuration to $SSH_CONFIG"
else
  echo "A github.com entry already exists in $SSH_CONFIG"
  echo "Leaving the existing entry unchanged."
fi

# Start an SSH agent if necessary.
eval "$(ssh-agent -s)" >/dev/null

# Add the key to the macOS keychain and SSH agent.
if ssh-add --apple-use-keychain "$KEY_PATH" 2>/dev/null; then
  echo "SSH key added to the macOS keychain."
else
  echo "The --apple-use-keychain option was unavailable."
  echo "Adding the key using the standard ssh-add command."
  ssh-add "$KEY_PATH"
fi

# Copy the public key to the clipboard.
pbcopy < "$PUBLIC_KEY_PATH"

echo
echo "Your public SSH key has been copied to the clipboard."
echo "GitHub's SSH key settings page will now open."
echo
echo "On GitHub:"
echo "  1. Select 'New SSH key'."
echo "  2. Use a title such as 'MDQXQPC75V Mac'."
echo "  3. Select 'Authentication Key'."
echo "  4. Paste the public key."
echo "  5. Select 'Add SSH key'."
echo

open "https://github.com/settings/ssh/new"

read -r -p "Press Enter after you have added the key to GitHub..."

echo
echo "Testing GitHub authentication..."

# GitHub's successful SSH test can return status 1 because shell access
# is deliberately disabled, so inspect its output rather than its status.
SSH_OUTPUT="$(ssh -T git@github.com 2>&1 || true)"
echo "$SSH_OUTPUT"

if [[ "$SSH_OUTPUT" != *"successfully authenticated"* ]]; then
  echo
  echo "ERROR: GitHub did not authenticate this SSH key."
  echo
  echo "Confirm that:"
  echo "  - You pasted the contents of $PUBLIC_KEY_PATH"
  echo "  - The key was added as an Authentication Key"
  echo "  - You added it to the correct GitHub account"
  exit 1
fi

echo
echo "GitHub SSH authentication succeeded."

# Confirm this is a Git repository.
if ! git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
  echo
  echo "ERROR: The current directory is not a Git repository."
  echo "Run this script from inside your extracted powerpoint directory."
  exit 1
fi

# Configure the GitHub remote.
if git remote get-url origin >/dev/null 2>&1; then
  git remote set-url origin "$REMOTE_URL"
else
  git remote add origin "$REMOTE_URL"
fi

echo
echo "Configured Git remote:"
git remote -v

# Ensure the local branch is named main.
git branch -M main

# Commit any uncommitted files.
if [[ -n "$(git status --porcelain)" ]]; then
  echo
  echo "Uncommitted changes found. Creating a commit..."
  git add --all
  git commit -m "Update Markdown to PowerPoint generator"
else
  echo
  echo "No uncommitted changes found."
fi

echo
echo "Pushing main to GitHub..."
git push -u origin main

echo
echo "Done."
echo "Repository:"
echo "  https://github.com/frankwoconnor/powerpoint"
