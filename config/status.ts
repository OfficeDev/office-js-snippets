import * as nodeStatus from 'node-status';
import { isString, find, isNil, isArray } from 'lodash';

interface IStage {
    steps: any[];
    count: number;
    doneStep: (completed: boolean, message?: string) => void;
};

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

        nodeStatus.start({ pattern: '{spinner.cyan} {job.step}' });
    }

    /**
     * Update the status stage
     * @param error Error object.
     */
    complete(success: boolean, stage: string, additionalDetails?: string | Error | Array<string | Error>) {
        const array = getDetailsArray();

        if (array.length === 0) {
            this.stages.doneStep(success);
        } else {
            // Add a newline before
            array.splice(0, 0, '');
            this.stages.doneStep(success, array.join('\n    * ') + '\n');
        }

        this.steps[stage] = false;

        if (!find(this.steps as any, (item) => item === true)) {
            nodeStatus.stop();
        }


        // Helper:
        function getDetailsArray() {
            if (!isArray(additionalDetails)) {
                additionalDetails = [additionalDetails];
            }

            return additionalDetails
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
                    return stringified;
                })
                .filter(item => !isNil(item));
        }
    };

    /**
     * Add a new stage and complete the previous stage.
     * @param stage Name of the stage.
     */
    add(stage: string) {
        this.stages.steps.push(stage);
        this.steps[stage] = true;
    };
};

export const status = new Status();
export const console = status.console;
