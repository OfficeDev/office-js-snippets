import * as nodeStatus from 'node-status';

interface IStage {
    steps: any[];
    count: number;
    doneStep: (completed: boolean, err: Error) => void;
};

export class Status {
    stages: IStage;
    active: boolean;

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
    complete(error?: Error): boolean {
        if (!this.active) {
            return false;
        }

        this.active = false;
        this.stages.doneStep(error == null, error);
        if (this.stages.count >= this.stages.steps.length) {
            nodeStatus.stop();
        }

        return true;
    };

    /**
     * Add a new stage and complete the previous stage.
     * @param stage Name of the stage.
     */
    add(stage: string): boolean {
        if (this.active) {
            return false;
        }
        this.stages.steps.push(stage);
        this.active = true;
        return true;
    };
};

export const status = new Status();
export const console = status.console;
