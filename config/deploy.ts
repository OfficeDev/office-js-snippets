#!/usr/bin/env node --harmony

import * as chalk from 'chalk';
import * as shell from 'shelljs';
import { isString } from 'lodash';
import { status } from './status';
import { banner, destinationBranch } from './helpers';

const { TRAVIS, TRAVIS_BRANCH, TRAVIS_PULL_REQUEST, GH_ACCOUNT, GH_TOKEN, GH_REPO, TRAVIS_COMMIT_MESSAGE } = process.env;

(() => {
    try {
        if (precheck()) {
            const URL = `https://${GH_TOKEN}@github.com/${GH_ACCOUNT}/${GH_REPO}.git`;
            status.add('Pushing to GitHub');
            deployBuild(URL);
            status.complete(true, 'Pushing to GitHub');
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
    if (!isString(GH_ACCOUNT)) {
        throw new Error('"AZURE_WA_USERNAME" is a required global variable.');
    }

    /* Check if the password is configured. If not abort immediately. */
    if (!isString(GH_TOKEN)) {
        throw new Error('"AZURE_WA_PASSWORD" is a required global variable.');
    }

    return true;
}

async function deployBuild(url) {
    const start = Date.now();
    shell.exec('git config --add user.name "Travis CI"');
    shell.exec('git config --add user.email "travis.ci@microsoft.com"');
    shell.exec('git checkout --orphan newbranch');
    shell.exec('git reset');

    let result: any = shell.exec('git add -f samples playlists');
    if (result.code !== 0) {
        shell.echo(result.stderr);
        throw new Error('An error occurred while adding files...');
    }

    result = shell.exec('git commit -m "' + TRAVIS_COMMIT_MESSAGE + '"');
    if (result.code !== 0) {
        shell.echo(result.stderr);
        throw new Error('An error occurred while commiting files...');
    }

    const gitPushCommand = `git push ${url} -q -f -u HEAD:refs/heads/${destinationBranch(TRAVIS_BRANCH)}`;
    console.log(chalk.bold.cyan('Pushing snippets & playlists... Please wait...'));
    result = shell.exec(gitPushCommand, { silent: true });
    if (result.code !== 0) {
        throw new Error(`An error occurred while executing ${gitPushCommand}`);
    }

    const end = Date.now();
    console.log(chalk.bold.cyan('Successfully deployed in ' + (end - start) / 1000 + ' seconds.', 'green'));
}
