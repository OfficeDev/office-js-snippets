const chalk = require('chalk');
const path = require('path');
const fs = require('fs');
const jsYaml = require('js-yaml');
const _ = require('lodash');

const walk = (dir, files = []) => {
    fs.readdirSync(dir).forEach(file =>
        files = fs.statSync(path.join(dir, file)).isDirectory() ?
            walk(path.join(dir, file), files) :
            files.concat(path.join(dir, file))
    );

    return files;
}

const stripNames = (file) => {
    let url = file;
    console.log(chalk.bold.blue(file));
    let strippedPath = file.replace(path.resolve('samples'), '').replace('\\', '');
    let [source, group, filename] = strippedPath.split('\\');
    return { source, group, url, filename };
}

const process = (file) => {
    let {source, group, url, filename} = stripNames(file);
    console.log(chalk.bold.green(`${source}\\${group}\\${filename}`));

    let contents = fs.readFileSync(url);
    let {name, description} = jsYaml.load(contents);
    let gist = `https://raw.githubusercontent.com/WrathOfZombies/samples/master/samples/${source}/${group}/${filename}`;
    group = _.startCase(group);
    return { source, meta: { name, description, gist, group } };
}

console.log(chalk.bold.yellow('Loading samples...'));
let files = walk(path.resolve('samples')).sort();

let processedFiles = files.map(process);
let groupedFiles = _.groupBy(processedFiles, 'source');

fs.mkdirSync(path.resolve('playlists'));

_.each(groupedFiles, (items, group) => {
    let pluckedItems = _.map(items, 'meta');
    let contents = jsYaml.safeDump(pluckedItems);
    fs.writeFileSync(path.resolve(`playlists/${group}.yaml`), contents);
    console.log(chalk.bold.yellow(`\nCreated ${group}.yaml`));
});