export interface IFormProps {
    id: string;
    name?: string;
    formName?: string;
    workflow?: {
        name?: string;
    };
    urls?: {
        formUrl?: string;
    };
    // Legacy fallbacks from docs example (not returned by actual API)
    url?: string;
    formUrl?: string;
    workflowName?: string;
    description?: string;
    status?: string;
}
