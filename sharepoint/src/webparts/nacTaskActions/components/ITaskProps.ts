export interface ITaskProps {
    id: string;
    name: string;
    description: string;
    dueDate: string;
    status: string;
    message?: string; // HTML email body of the task
    outcomes?: string[]; // Alias — API returns this as "validOutcomes"
    validOutcomes?: string[]; // Correct API field name per uiRequest=true response
    expressApproval?: string; // "include" means express approval is enabled for this task
    taskAssignments: ITaskAssignment[];
    // Add further properties here if you want to capture more information about the tasks
}

export interface ITaskAssignment {
    id: string;
    status: string;
    assignee?: string;
    urls?: {
        formUrl?: string;
    }
    // Add further properties here if you want to capture more information about the task assignments
}
