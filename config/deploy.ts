#!/usr/bin/env node --harmony

import * as chalk from 'chalk';
import * as simpleGit from 'simple-git';
import { isString } from 'lodash';
import { status } from './status';
import { banner } from './helpers';

const git = simpleGit();
const { TRAVIS, TRAVIS_BRANCH, TRAVIS_PULL_REQUEST, GH_ACCOUNT, GH_TOKEN, GH_REPO, TRAVIS_COMMIT_MESSAGE } = process.env;

let url = `https://${GH_TOKEN}@github.com/${GH_ACCOUNT}/${GH_REPO}.git`;

(async () => {
    try {
        if (precheck()) {
            /* Pushing to GitHub */
            status.add('Pushing to GitHub');
            await deployBuild(url);
            status.complete('Pushing to GitHub');
        }
    }
    catch (exception) {
        handleError(exception);
    }
})();

/**
 * Deploying to GitHub
 */
async function deployBuild(url) {
    return new Promise((resolve, reject) => {
        const start = Date.now();
        try {
            git.addConfig('user.name', 'Travis CI')
                .addConfig('user.email', 'travis.ci@microsoft.com')
                .checkout(['--orphan', 'prod'])
                .add(['samples/', '-f'], (err) => {
                    if (err) {
                        return reject(err.replace(url, ''));
                    }
                })
                .add(['playlists/', '-f'], (err) => {
                    if (err) {
                        return reject(err.replace(url, ''));
                    }
                })
                .commit(TRAVIS_COMMIT_MESSAGE, () => console.log(chalk.bold.cyan('Pushing snippets & playlists... Please wait...')))
                .push(['-f', '-u', url, 'HEAD:refs/heads/prod'], (err) => {
                    if (err) {
                        return reject(err.replace(url, ''));
                    }

                    const end = Date.now();
                    console.log(chalk.bold.cyan('Successfully deployed in ' + (end - start) / 1000 + ' seconds.', 'green'));
                    return resolve();
                });
        }
        catch (error) {
            return reject(error);
        }
    });
}

function precheck() {
    /* Check if the code is running inside of travis.ci. If not abort immediately. */
    if (!TRAVIS) {
        console.log('Not running inside of Travis. Skipping deploy.');
        return false;
    }

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
