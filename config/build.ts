#!/usr/bin/env node --harmony

import * as path from 'path';
import * as chalk from 'chalk';
import { status } from './status';
import * as simpleGit from 'simple-git';
import { rmRf, mkDir, getFiles, writeFile, File, banner, loadYamlFile } from './helpers';
import { isString, startCase, groupBy, map } from 'lodash';
import * as jsyaml from 'js-yaml';
import 'rxjs/add/operator/map';
import 'rxjs/add/operator/filter';

const git = simpleGit();

const { TRAVIS, TRAVIS_BRANCH, TRAVIS_PULL_REQUEST, GH_ACCOUNT, GH_TOKEN, GH_REPO, TRAVIS_COMMIT_MESSAGE } = process.env;
const files: File[] = [];

(async () => {
    try {
        /* Creating playlists directory */
        status.add('Creating \'playlists\' folder');
        await rmRf('playlists');
        await mkDir('playlists');
        status.complete('Creating \'playlists\' folder');

        /* Loading samples */
        status.add('Loading snippets');
        let files$ = getFiles(path.resolve('samples'), path.resolve('samples'));
        status.complete('Loading snippets');

        files$.mergeMap(async (file) => {
            try {
                status.add(`Processing ${file.host}::${file.name}`);
                let { name, description } = await loadYamlFile<{ name: string, description: string }>(path.resolve('samples', file.path));
                status.complete(`Processing ${file.host}::${file.name}`);
                return {
                    name,
                    description,
                    ...file,
                    gist: `https://raw.githubusercontent.com/${GH_ACCOUNT}/${GH_REPO}/deployment/samples/${file.host}/${file.group}/${file.name}`,
                    group: startCase(file.group)
                };
            } catch (exception) {
                status.complete(`Processing ${file.host}::${file.name}`, exception);
                handleError(`Failed to process ${file.host}::${file.name}: ${exception.message || exception}`);
                return null;
            }
        })
            .filter((file) => !(file == null) && file.name !== 'default.yaml')
            .map((file) => files.push(file))
            .subscribe(null, handleError, snippetsProcessed);
    }
    catch (exception) {
        handleError(exception);
    }
})();

/**
 * Generic error handler.
 * @param error Error object.
 */
function handleError(error?: any) {
    banner('An error has occured', error.message || error, chalk.bold.red);
    process.exit(1);
};

/**
 * Generating playlists
 */
async function snippetsProcessed() {
    if (files == null || files.length < 1) {
        return;
    }

    /* Generating playlists */
    status.add('Generating playlists');
    const groups = groupBy(files, 'host');
    let promises = map(groups, async (items, group) => {
        let contents = jsyaml.safeDump(items);
        await writeFile(path.resolve(`playlists/${group}.yaml`), contents);
        banner(`Created ${group}.yaml`);
    });
    await Promise.all(promises);
    status.complete('Generating playlists');

    let url = `https://${GH_TOKEN}@github.com/${GH_ACCOUNT}/${GH_REPO}.git`;

    if (precheck()) {
        /* Pushing to GitHub */
        status.add('Pushing to GitHub');
        await deployBuild(url);
        status.complete('Pushing to GitHub');
    }
};

/**
 * Deploying to GitHub
 */
async function deployBuild(url) {
    return new Promise((resolve, reject) => {
        const start = Date.now();
        try {
            git.addConfig('user.name', 'Travis CI')
                .addConfig('user.email', 'travis.ci@microsoft.com')
                .checkout('HEAD')
                .add(['samples/', '-A', '-f'], (err) => {
                    if (err) {
                        return reject(err.replace(url, ''));
                    }
                })
                .add(['playlists/', '-A', '-f'], (err) => {
                    if (err) {
                        return reject(err.replace(url, ''));
                    }
                })
                .commit(TRAVIS_COMMIT_MESSAGE, () => console.log(chalk.bold.cyan('Pushing ' + path + '... Please wait...')))
                .push(['-f', '-u', url, 'HEAD:refs/heads/deployment'], (err) => {
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
