import * as path from 'path';
import * as fs from 'fs';
import * as jsYaml from 'js-yaml';
import * as chalk from 'chalk';

const stripNames = (file) => {
    let url = file;
    let strippedPath = file.replace(path.resolve('samples'), '').replace('/', '');
    let [source, group, filename] = strippedPath.split('/');
    return { source, group, url, filename };
};

const process = (file) => {
    let { source, group, url, filename } = stripNames(file);
    console.log(chalk.bold.green(`${source}/${group}/${filename}`));

    let contents = fs.readFileSync(url);
    let { name, description } = jsYaml.load(contents);
    let gist = `https://raw.githubusercontent.com/WrathOfZombies/samples/master/samples/${source}/${group}/${filename}`;
    group = _.startCase(group);
    return { source, meta: { name, description, gist, group } };
}

console.log(chalk.bold.yellow('Loading samples...'));
let files = walk(path.resolve('samples')).sort();
files = files.filter(file => !/default.yaml?$/i.test(file));
let processedFiles = files.map(process);
let groupedFiles = _.groupBy(processedFiles, 'source');

fs.mkdirSync(path.resolve('playlists'));

_.each(groupedFiles, (items, group) => {
    let pluckedItems = _.map(items, 'meta');
    let contents = jsYaml.safeDump(pluckedItems);
    fs.writeFileSync(path.resolve(`playlists/${group}.yaml`), contents);
    console.log(chalk.bold.yellow(`\nCreated ${group}.yaml`));
});
