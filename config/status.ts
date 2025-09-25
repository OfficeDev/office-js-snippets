import chalk from 'chalk';
import { isString, isNil, isArray } from 'lodash';

interface IStage {
    steps: any[];
    count: number;
    doneStep: (completed: boolean, message?: string) => void;
}

export class Status {
    stages: IStage;
    steps: { [step: string]: boolean } = {};

    get console() {
        // Return the global console object methods
        return {
            log: global.console.log.bind(global.console),
            error: global.console.error.bind(global.console),
            warn: global.console.warn.bind(global.console),
            info: global.console.info.bind(global.console)
        };
    }

    constructor() {
        /* Initialize the simple status system */
        this.stages = {
            steps: [],
            count: 0,
            doneStep: this.doneStep.bind(this)
        };
    }

    private doneStep(completed: boolean, message?: string): void {
        const symbol = completed ? chalk.green('✓') : chalk.red('✗');
        if (message) {
            global.console.log(`${symbol} ${message}`);
        }
    }

    /**
     * Update the status stage
     * @param error Error object.
     */
    complete(success: boolean, stage: string, additionalDetails?: string | Error | Array<string | Error>) {
        if (!isArray(additionalDetails)) {
            additionalDetails = [additionalDetails];
        }

        success = success && additionalDetails.findIndex(item => item instanceof Error) < 0;

        const messageArray = getDetailsArray();
        const symbol = success ? chalk.green('✓') : chalk.red('✗');

        // Log the completion with symbol
        global.console.log(`${symbol} ${stage}`);

        if (messageArray.length > 0) {
            messageArray.forEach(msg => {
                global.console.log(`    * ${msg}`);
            });
        }

        this.steps[stage] = false;

        // Helper:
        function getDetailsArray() {
            return (additionalDetails as any[])
                .map(item => {
                    if (isNil(item)) {
                        return null;
                    }

                    if (isString(item)) {
                        return item;
                    }

                    let stringified = item.message || item.toString();
                    if (stringified === '[object Object]') {
                        stringified = JSON.stringify(item);
                    }

                    if (item instanceof Error) {
                        return chalk.bold.red(stringified);
                    } else {
                        return stringified;
                    }
                })
                .filter(item => !isNil(item));
        }
    }

    /**
     * Add a new stage and mark it as started.
     * @param stage Name of the stage.
     */
    add(stage: string) {
        this.stages.steps.push(stage);
        this.steps[stage] = true;
        global.console.log(chalk.cyan(`○ ${stage}`));
    }
}

export const status = new Status();
export const console = status.console;
