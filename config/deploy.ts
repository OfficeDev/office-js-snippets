#!/usr/bin/env node --harmony

import * as chalk from 'chalk';
import * as shell from 'shelljs';
import { isString } from 'lodash';
import { status } from './status';
import { banner } from './helpers';

const { TRAVIS, TRAVIS_BRANCH, TRAVIS_PULL_REQUEST, GH_ACCOUNT, GH_TOKEN, GH_REPO, TRAVIS_COMMIT_MESSAGE } = process.env;

try {
    if (precheck()) {
        const URL = `https://${GH_TOKEN}@github.com/${GH_ACCOUNT}/${GH_REPO}.git`;
        status.add('Pushing to GitHub');
        deployBuild(URL);
        status.complete(true, 'Pushing to GitHub');
    }
}
catch (exception) {
    handleError(exception);
}

/**
 * Deploying to GitHub
 */
async function deployBuild(url) {
    try {
        const start = Date.now();
        shell.exec('git config --add user.name "Travis CI"');
        shell.exec('git config --add user.email "travis.ci@microsoft.com"');
        shell.exec('git checkout --orphan newbranch');
        shell.exec('git reset');
        let result: any = shell.exec('git add -f samples playlists');
        if (result.code !== 0) {
            shell.echo(result.stderr);
            handleError('An error occurred while adding files...');
        }
        result = shell.exec('git commit -m "' + TRAVIS_COMMIT_MESSAGE + '"');
        if (result.code !== 0) {
            shell.echo(result.stderr);
            handleError('An error occurred while commiting files...');
        }
        console.log(chalk.bold.cyan('Pushing snippets & playlists... Please wait...'));
        result = shell.exec('git push ' + url + ' -q -f -u HEAD:refs/heads/prod', { silent: true });
        if (result.code !== 0) {
            handleError('An error occurred while deploying playlists to ...');
        }
        const end = Date.now();
        console.log(chalk.bold.cyan('Successfully deployed in ' + (end - start) / 1000 + ' seconds.', 'green'));
    }
    catch (error) {
        handleError('Deployment failed...');
    }
}

function precheck() {
    /* Check if the code is running inside of travis.ci. If not abort immediately. */
    if (!TRAVIS) {
        console.log('Not running inside of Travis. Skipping deploy.');
        return false;
    }

    // Careful! Need this check because otherwise, a pull request against master would immediately trigger a deployment.
    // On the other hand, TODO (issue #6): Still need to make it so that a MERGED pull request INTO master *does* create a deployment.
    if (TRAVIS_PULL_REQUEST !== 'false') {
        console.log('Skipping deploy for pull requests.');
        return false;
    }

    if (TRAVIS_BRANCH !== 'master') {
        console.log('Skipping deploy for non master branches.');
        return false;
    }

    /* Check if the username is configured. If not abort immediately. */
    if (!isString(GH_ACCOUNT)) {
        handleError('"AZURE_WA_USERNAME" is a required global variable.');
    }

    /* Check if the password is configured. If not abort immediately. */
    if (!isString(GH_TOKEN)) {
        handleError('"AZURE_WA_PASSWORD" is a required global variable.');
    }

    return true;
}

/**
 * Generic error handler.
 * @param error Error object.
 */
function handleError(error?: any) {
    banner('An error has occured', error.message || error, chalk.bold.red);
    process.exit(1);
}
