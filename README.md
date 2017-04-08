[![Build Status](https://travis-ci.com/OfficeDev/script-lab-samples.svg?token=zKp5xy2SuSortMzv5Pqc&branch=master)](https://travis-ci.com/OfficeDev/script-lab-samples)

# Script Lab Samples
A collection of samples for Script Lab

## To contribute:
- Fork this repo
- Add samples using the instructions below
- Submit a pull request that one of the repo contributors will review and merge in.

## Folder Structure
- All snippets must be inside the samples folder.
- The `base folders` such as Excel, Word etc. are all the various broad level categories.
- Inside of each `base folder`, there are `group folders` for the group in which a sample belongs to.
- Inside of each `group folder`, there are `.yaml` which represent a snippet.

## Adding a new sample

1. Create a snippet using Script Lab.
2. Click on `Copy to Clipboard` in the `Share` menu.
3. Fill in the `author`, `name`, `description`, `source` properties if they are empty so that we can generate the playlist correctly.
4. Add that snippet into the respective folders. Make sure that the snippet file names and folder names are in [`kebab-case`](http://wiki.c2.com/?KebabCase).

## Style guidelines:

Basic snippet structure is as follows:

    $("#run").click(run);

    async function run() {
        try {
            await Word.run(async (context) => {
                let range = context.document.getSelection();
                range.font.color = "red";

                await context.sync();
            });
        }
        catch (error) {
            OfficeHelpers.Utilities.log(error);
        }
    }

A few style rules to observe:

* Each button-click handler should have its own `async` function, called "run" if there is only one button on the page -- otherwise, name it as you will.
* Inside the function there shall be a try/catch.  In it you will await the `Excel.run` or `Word.run`, and use `async/await` inside of the `.run` as well.
* All HTML IDs should be `all-lower-case-and-hyphenated`.
* Unless you are explicitly showing pretty UI, I wouldn't do the popup notification except for one or two samples.  It's a lot of HTML & JS code, and it's also not strictly Fabric-y (there is a more "correct" way of doing this with components).
* Strings should be in double-quotes.
* Don't forget the semicolons.
* `Libraries` in snippets must have a specific version. Eg. `jquery@3.1.1`.
