#!/usr/bin/env node --harmony

import * as chalk from 'chalk';
import * as shell from 'shelljs';
import { forIn } from 'lodash';
import { isString } from 'lodash';
import { banner, getDestinationBranch } from './helpers';

interface IEnvironmentVariables {
    TRAVIS: string,
    TRAVIS_BRANCH: string,
    TRAVIS_PULL_REQUEST: string,
    TRAVIS_COMMIT_MESSAGE: string,
    GH_ACCOUNT: string,
    GH_REPO: string,
    GH_TOKEN: string
}

const environmentVariables: IEnvironmentVariables = process.env as any;

(() => {
    try {
        // Note, if precheck fails, it will do its own banner, so only need to focus on the true case.
        if (precheck()) {
            const destinationBranch = getDestinationBranch(environmentVariables.TRAVIS_BRANCH);
            const repoUrl = `https://github.com/${environmentVariables.GH_ACCOUNT}/${environmentVariables.GH_REPO}/tree/${destinationBranch}`;
            banner('Starting deployment', repoUrl);

            const start = Date.now();
            shell.exec('git config --add user.name "Travis CI"');
            shell.exec('git config --add user.email "travis.ci@microsoft.com"');
            shell.exec('git checkout --orphan newbranch');
            shell.exec('git reset');

            execCommand('git add -f samples private-samples playlists view snippet-extractor-output README.md');
            execCommand(`git commit -m "Travis auto-deploy of ${environmentVariables.TRAVIS_COMMIT_MESSAGE.replace(/\W/g, '_')} [skip ci]"`);

            const tokenizedGitHubGitUrl = `https://<<<token>>>@github.com/${environmentVariables.GH_ACCOUNT}/${environmentVariables.GH_REPO}.git`;
            execCommand(`git push ${tokenizedGitHubGitUrl} -f -u HEAD:refs/heads/${destinationBranch}`, {
                token: environmentVariables.GH_TOKEN
            });

            const end = Date.now();

            banner('Deployment succeeded', `Successfully deployed to ${repoUrl} in ${(end - start) / 1000} seconds.`, chalk.bold.green);
        }
    }
    catch (error) {
        banner('An error has occurred', error.message || error, chalk.bold.red);
        banner('DEPLOYMENT DID NOT GET TRIGGERED', error.message || error, chalk.bold.red);

        // Even though deployment failure does not imply dev failure, we want to break the build
        // to make it obvious that the deployment went wrong
        process.exit(1);
    }

    process.exit(0);
})();


function precheck() {
    /* Check if the code is running inside of travis.ci. If not abort immediately. */
    if (!environmentVariables.TRAVIS) {
        banner('Deployment skipped', 'Not running inside of Travis.', chalk.yellow.bold);
        return false;
    }

    // Careful! Need this check because otherwise, a pull request against master would immediately trigger a deployment.
    if (environmentVariables.TRAVIS_PULL_REQUEST !== 'false') {
        banner('Deployment skipped', 'Skipping deploy for pull requests.', chalk.yellow.bold);
        return false;
    }

    if (getDestinationBranch(environmentVariables.TRAVIS_BRANCH) == null) {
        banner('Deployment skipped', 'Skipping deploy for pull requests.', chalk.yellow.bold);
        return false;
    }

    /* Check if the username is configured. If not abort immediately. */
    const requiredFields: Array<keyof IEnvironmentVariables> = ['GH_ACCOUNT', 'GH_REPO', 'GH_TOKEN'];
    requiredFields.forEach(key => {
        if (!isString(environmentVariables[key])) {
            throw new Error(`"${key}" is a required global variables.`);
        }
    });

    return true;
}

/**
 * Execute a shell command.
 * @param originalSanitizedCommand - The command to execute. Note that if it contains something secret, put it in triple <<<NAME>>> syntax, as the command itself will get echo-ed.
 * @param secretSubstitutions - key-value pairs to substitute into the command when executing.  Having any secret substitutions will automatically make the command run silently.
 */
function execCommand(originalSanitizedCommand: string, secretSubstitutions = {}) {
    console.log(originalSanitizedCommand);

    let hadSecrets = false;
    let command = originalSanitizedCommand;
    forIn(secretSubstitutions, (value, key) => {
        hadSecrets = true;
        command = replaceAll(command, '<<<' + key + '>>>', value);
    });

    if (hadSecrets) {
        console.log(chalk.yellow('Command contained secret substitution values; running the `shell.exec` silently'));
    }

    let result: any = shell.exec(command, hadSecrets ? { silent: true } : null);
    if (result.code !== 0) {
        throw new Error(`An error occurred while executing "${originalSanitizedCommand}"`);
    }
}

function replaceAll(source, search, replacement) {
    return source.split(search).join(replacement);
}
