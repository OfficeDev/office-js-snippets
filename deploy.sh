#!/bin/bash
set -e # Exit with nonzero exit code if anything fails

# Inspired by https://gist.github.com/domenic/ec8b0fc8ab45f39403dd

BRANCH="master"

# Only commits to master branch will trigger a build.
if [ "$TRAVIS_BRANCH" != "$BRANCH" ]; then
    echo "Skipping deploy."
    exit 0
fi

# Save some useful information
REPO=`git config remote.origin.url`
SHA=`git rev-parse --verify HEAD`

# Clone the existing repo into `out`and cd into it
git clone $REPO out
cd out
git checkout -b deployment

# Run `npm install` and our `build` script
npm install
npm run build
git status

# Now let's go have some fun with the cloned repo
git config user.name "Travis CI"
git config user.email "$COMMIT_AUTHOR_EMAIL"

# Commit the "changes", i.e. the new version.
# The delta will show diffs between new and old versions.
git add .
git commit -m "Travis: auto-generating playlists [${SHA}]"

# Get the deploy key by using Travis's stored variables to decrypt deploy_key.enc
ENCRYPTED_KEY_VAR="encrypted_${ENCRYPTION_LABEL}_key"
ENCRYPTED_IV_VAR="encrypted_${ENCRYPTION_LABEL}_iv"
ENCRYPTED_KEY=${!ENCRYPTED_KEY_VAR}
ENCRYPTED_IV=${!ENCRYPTED_IV_VAR}
openssl aes-256-cbc -K $ENCRYPTED_KEY -iv $ENCRYPTED_IV -in deploy_key.enc -out deploy_key -d
chmod 600 deploy_key
eval `ssh-agent -s`
ssh-add deploy_key

# Now that we're all set up, we can push.
git push -u origin deployment