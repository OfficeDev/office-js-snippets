# Scripts

These scripts used to help maintain this repository.

## Setup

> npm install

## Edit Script

This script is used to bulk edit samples.

To run this script:

> npm run edit

The edit targets all prod samples listed in `playlists-prod` and all default samples.

Under the src folder the transform* files contain the specific JavaScript transforms that will run.

To develop new transforms:

1. Make changes the transform* functions
2. Run the transforms (npm run edit)
3. Check using the git diff to make sure the changes are what you expect
4. If you don't like the changes run the following in the **samples** folder:
    > git checkout -- *
