import * as nodeStatus from 'node-status';
import { find } from 'lodash';

interface IStage {
    steps: any[];
    count: number;
    doneStep: (completed: boolean, err: Error) => void;
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
    complete(stage: string, error?: Error) {
        this.stages.doneStep(error == null, error);
        this.steps[stage] = false;

        if (!find(this.steps as any, (item) => item === true)) {
            nodeStatus.stop();
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
