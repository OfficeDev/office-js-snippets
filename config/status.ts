import * as nodeStatus from 'node-status';
import chalk from 'chalk';
import { isString, find, isNil, isArray } from 'lodash';

interface IStage {
    steps: any[];
    count: number;
    doneStep: (completed: boolean, message?: string) => void;
}

export class Status {
    stages: IStage;
    steps: { [step: string]: boolean } = {};

    get console() {
        return nodeStatus.console();
    }

    constructor() {
        /* Initialize the status library */
        this.stages = nodeStatus.addItem('stages', {
            steps: []
        });

        nodeStatus.start();
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

        if (messageArray.length === 0) {
            this.stages.doneStep(success);
        } else {
            // Add a newline before
            messageArray.splice(0, 0, '');
            this.stages.doneStep(success, messageArray.join('\n    * ') + '\n');
            //FIXME `${chalk.bold.red('WARNING: one of the messages above was an error')}`)
        }

        this.steps[stage] = false;

        if (!find(this.steps as any, (item) => item === true)) {
            nodeStatus.stop();
        }


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
     * Add a new stage and complete the previous stage.
     * @param stage Name of the stage.
     */
    add(stage: string) {
        this.stages.steps.push(stage);
        this.steps[stage] = true;
    }
}

export const status = new Status();
export const console = status.console;
