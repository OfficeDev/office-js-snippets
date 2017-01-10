[![Build Status](https://travis-ci.org/WrathOfZombies/samples.svg?branch=master)](https://travis-ci.org/WrathOfZombies/samples)

# Add-in Playground Samples
A collection of samples for the Add-in Playground

## Folder Structure
- All snippets must be inside the samples folder.
- The `base folders` such as Excel, Word etc. are all the various broad level categories.
- Inside of each `base folder`, there are `group folders` for the group in which a sample belongs to.
- Inside of each `group folder`, there are `.yaml` which represent a snippet.

## Adding a new sample

1. Create a snippet using the Add-in Playground.
2. Click on `Copy to Clipboard` in the `Share` menu.
3. Fill in the `author`, `name`, `description`, `source` properties if they are empty so that we can generate the playlist correctly.
4. Add that snippet into the respective folders. Make sure that the snippet file names and folder names are in `kebabcase`.