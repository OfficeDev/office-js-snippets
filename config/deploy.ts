#!/usr/bin/env node --harmony

import * as chalk from 'chalk';
import * as shell from 'shelljs';
import { forIn } from 'lodash';
import { isString } from 'lodash';
import { status } from './status';
import { banner, destinationBranch } from './helpers';

const { TRAVIS, TRAVIS_BRANCH, TRAVIS_PULL_REQUEST, GH_ACCOUNT, GH_TOKEN, GH_REPO, TRAVIS_COMMIT_MESSAGE } = process.env;

(() => {
    try {
        if (precheck()) {
            deployBuild(URL);
        } else {
            console.log('Did not pass pre-check. Exiting.');
            process.exit(0);
        }
    }
    catch (error) {
        banner('An error has occured', error.message || error, chalk.bold.red);
        process.exit(1);
    }
})();


function precheck() {
    /* Check if the code is running inside of travis.ci. If not abort immediately. */
    if (!TRAVIS) {
        console.log('Not running inside of Travis. Skipping deploy.');
        return false;
    }

    // Careful! Need this check because otherwise, a pull request against master would immediately trigger a deployment.
    if (TRAVIS_PULL_REQUEST !== 'false') {
        console.log('Skipping deploy for pull requests.');
        return false;
    }

    if (destinationBranch(TRAVIS_BRANCH) == null) {
        console.log('Skipping deploy for non `master` or `prod` branches.');
        return false;
    }

    /* Check if the username is configured. If not abort immediately. */
    if (!isString(GH_ACCOUNT) || !isString(GH_TOKEN)) {
        throw new Error('"GH_ACCOUNT" and "GH_TOKEN" are required global variables.');
    }

    return true;
}

async function deployBuild(url) {
    status.add('Pushing to GitHub');

    const start = Date.now();
    shell.exec('git config --add user.name "Travis CI"');
    shell.exec('git config --add user.email "travis.ci@microsoft.com"');
    shell.exec('git checkout --orphan newbranch');
    shell.exec('git reset');

    execCommand('git add -f samples playlists');
    execCommand('git commit -m "' + TRAVIS_COMMIT_MESSAGE + '"');

    execCommand(`git push <<<url>>> -q -f -u HEAD:refs/heads/${destinationBranch(TRAVIS_BRANCH)}`, {
        url: `https://${GH_TOKEN}@github.com/${GH_ACCOUNT}/${GH_REPO}.git`
    });

    const end = Date.now();
    status.complete(true, 'Pushing to GitHub', chalk.bold.green(`Successfully deployed in ${(end - start) / 1000} seconds.`));
}

/**
 * Execute a shall command.
 * @param command - The command to execute. Note that if it contains something secret, put it in tripple <<<NAME>>> syntax, as the command itself will get echo-ed.
 * @param secretSubstitutions - key-value pairs to substitute into the command when executing.
 */
function execCommand(command: string, secretSubstitutions = {}) {
    console.log(command);

    forIn(secretSubstitutions, (value, key) => command = replaceAll(command, '<<<' + key + '>>>', value));
    let result: any = shell.exec(command);
    if (result.code !== 0) {
        shell.echo(result.stderr);
        throw new Error(`An error occurred while executing "${command}"`);
    }
}

function replaceAll(source, search, replacement) {
    return source.split(search).join(replacement);
}
