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
SHA=`git rev-parse --verify HEAD`

# Now let's go have some fun with the cloned repo
git config --global user.name "Travis CI"
git config --global user.email "$COMMIT_AUTHOR_EMAIL"

# Clone the existing repo into `out`and cd into it
git clone "https://${GH_TOKEN}@github.com/WrathOfZombies/samples.git" out
cd out
git checkout -b deployment

# Run `npm install` and our `build` script
npm install
npm run build
git status

# Commit the "changes", i.e. the new version.
# The delta will show diffs between new and old versions.
git add .
git commit -m "Travis: auto-generating playlists [${SHA}]"

# Now that we're all set up, we can push.
git push -u origin deployment --force
exit 0