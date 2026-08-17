import * as React from 'react';
import styles from './NacTaskActions.module.scss';
import type { INacTaskActionsProps } from './INacTaskActionsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { AuthService } from './AuthService';
import axios from 'axios';
import { ITaskProps, ITaskAssignment } from './ITaskProps';
import { Config } from './Config';
import DOMPurify from 'dompurify';
import nintexColorLogo from '../assets/nintex_RGB_color_500.png';
import nintexReversedLogo from '../assets/nintex_RGB_reversed_500.png';

type StatusFilter = 'active' | 'expired' | 'complete' | 'overridden' | 'terminated' | 'all';

const STATUS_OPTIONS: StatusFilter[] = ['active', 'expired', 'complete', 'overridden', 'terminated', 'all'];

// Looks up a status-specific CSS module class by key, falling back to an
// empty string if no matching class exists for that status.
function getStatusClass(status: string): string {
  const key = `status-${status.toLowerCase()}`;
  const dict = styles as unknown as Record<string, string>;
  return dict[key] || '';
}

// Tracks which task/assignment + outcome is pending confirmation before
// the complete-task call is actually made (completion is irreversible).
interface IPendingAction {
  taskId: string;
  taskName: string;
  assignmentId: string;
  outcome: string;
  comment: string;
}

interface INacTaskActionsState {
  accessToken?: string;
  tasks?: ITaskProps[];
  statusFilter: StatusFilter;
  pendingAction?: IPendingAction;
  actionInFlight: boolean;
  actionError?: string;
  actionSuccessMessage?: string;
}

export default class NacTaskActions extends React.Component<INacTaskActionsProps, INacTaskActionsState, {}> {
  private authService: AuthService;
  private authConfig: Config;

  constructor(props: INacTaskActionsProps) {
    super(props);
    this.authConfig = new Config(this.props.clientId, this.props.redirectUri, this.props.region, this.props.useProd);
    this.authService = new AuthService(this.authConfig);
    this.state = {
      accessToken: undefined,
      tasks: undefined,
      statusFilter: 'active',
      pendingAction: undefined,
      actionInFlight: false,
      actionError: undefined,
      actionSuccessMessage: undefined
    };

    // If the URL contains ?code= this page load is the OAuth callback —
    // the redirect back from the Nintex auth server. Complete the token
    // exchange and fetch tasks.
    if (AuthService.isAuthCallback()) {
      this.finishAuthAndFetchTasks().catch((error) => { console.error(error); });
    }
  }

  // On mount, check localStorage for a valid (non-expired) token from a
  // previous login.  If found, load tasks immediately — no "Log In" click
  // needed.  Skipped when this page load is an OAuth callback (?code=),
  // because finishAuthAndFetchTasks() already handles that path.
  public componentDidMount(): void {
    if (!AuthService.isAuthCallback() && !this.state.accessToken) {
      this.tryExistingSession().catch((error) => {
        console.log('[NacTaskActions] No existing session found');
      });
    }
  }

  // If SPFx configuration properties are updated, recreate config and auth service
  public componentDidUpdate(prevProps: INacTaskActionsProps): void {
    if (prevProps.clientId !== this.props.clientId ||
        prevProps.redirectUri !== this.props.redirectUri ||
        prevProps.region !== this.props.region ||
        prevProps.useProd !== this.props.useProd) {
      this.authConfig = new Config(this.props.clientId, this.props.redirectUri, this.props.region, this.props.useProd);
      this.authService = new AuthService(this.authConfig);
    }
  }

  // Login button handler — full-page redirect to the Nintex auth server.
  private loginAndGetTasksClick = (): void => {
    this.authService.login().catch((error) => {
      console.error('[NacTaskActions] Login error:', error);
    });
  }

  // Completes the plain SharePoint redirect flow (called when ?code= is in URL
  // and we are NOT in Teams context).
  //
  // When two webparts share the same page, the first one to call
  // signinCallback() consumes the auth code and state entry from
  // localStorage.  The second one will get "No matching state found in
  // storage".  That is expected — the token is already in localStorage
  // from the first webpart, so we just fall back to reading it.
  private finishAuthAndFetchTasks = async (): Promise<void> => {
    try {
      await this.authService.handleAuthentication();
    } catch (error) {
      // Another webpart on this page likely already consumed the auth
      // code — the token exchange still happened, so the token should
      // be in localStorage.
      console.log('[NacTaskActions] signinCallback skipped (already handled):', (error as Error).message);
    }

    // Whether we handled the callback or the other webpart did, the
    // token should now be in localStorage.  Poll briefly in case the
    // other webpart's exchange is still in-flight.
    try {
      let accessToken = await this.authService.getAccessToken();
      if (!accessToken) {
        // Give the other webpart up to ~1.5 s to finish the exchange.
        for (let i = 0; i < 3 && !accessToken; i++) {
          await new Promise(resolve => setTimeout(resolve, 500));
          accessToken = await this.authService.getAccessToken();
        }
      }

      if (accessToken) {
        this.setState({ accessToken });
        this.getTasks(this.state.statusFilter).catch((error) => { console.error(error); });
      } else {
        console.error('[NacTaskActions] Access token missing after redirect callback');
      }
    } catch (error) {
      console.error('[NacTaskActions] Error during token retrieval', error);
    }
  }

  // Checks localStorage for a still-valid token from a previous login and
  // resumes the session without a redirect.  If the stored token is expired
  // but the user has an active Nintex session, attempts a silent renewal
  // via a hidden iframe.
  private tryExistingSession = async (): Promise<void> => {
    // 1. Check for a non-expired token in localStorage.
    let accessToken = await this.authService.getValidAccessToken();

    // 2. If no valid token, try silent renewal (hidden iframe).
    //    This works when the user has an active session on auth.nintexcloud.com.
    if (!accessToken) {
      console.log('[NacTaskActions] No valid stored token — trying silent login…');
      accessToken = await this.authService.trySilentLogin();
    }

    // 3. If we got a token (either stored or silently renewed), load data.
    if (accessToken) {
      console.log('[NacTaskActions] Resuming session (no login click needed)');
      this.setState({ accessToken });
      this.getTasks(this.state.statusFilter).catch((error) => { console.error(error); });
    }
    // Otherwise: no stored token, no active Nintex session → user sees
    // the "Log In" button (current default behavior).
  }

  private getTasks = async (status: StatusFilter): Promise<void> => {
    const region = this.authConfig.getRegion();
    const useProd = this.authConfig.getUseProd();
    const url = `https://${region}.nintex${useProd ? '' : 'test'}.io/workflows/v2/tasks?status=${status}&uiRequest=true`;

    try {
      if (!this.state.accessToken) {
        console.error('[NacTaskActions] Access token missing');
        return;
      }

      const response = await axios.get(url, {
        headers: {
          'Accept': 'application/json',
          'Authorization': `Bearer ${this.state.accessToken}`
        }
      });

      this.setState({ tasks: response.data.tasks });
    }
    catch (error) {
      if (axios.isAxiosError(error)) {
        console.error('[NacTaskActions] Get tasks error:', error.response ? error.response.data : error.message);
      } else {
        console.error('[NacTaskActions] Get tasks error:', error);
      }
    }
  }

  private handleFilterChange = (status: StatusFilter): void => {
    this.setState({ statusFilter: status });
    this.getTasks(status).catch((error) => { console.error(error); });
  }

  private handleOutcomeClick = (taskId: string, taskName: string, assignmentId: string | undefined, outcome: string): void => {
    if (!assignmentId) {
      this.setState({ actionError: 'No task assignment found for this task — cannot complete it.' });
      return;
    }
    this.setState({
      pendingAction: { taskId, taskName, assignmentId, outcome, comment: '' },
      actionError: undefined
    });
  }

  private handleCancelAction = (): void => {
    this.setState({ pendingAction: undefined });
  }

  private handleCommentChange = (comment: string): void => {
    this.setState((prevState) => ({
      pendingAction: prevState.pendingAction ? { ...prevState.pendingAction, comment } : undefined
    }));
  }

  private handleConfirmAction = async (): Promise<void> => {
    const { pendingAction, accessToken } = this.state;
    if (!pendingAction || !accessToken) {
      return;
    }

    const region = this.authConfig.getRegion();
    const useProd = this.authConfig.getUseProd();
    const url = `https://${region}.nintex${useProd ? '' : 'test'}.io/workflows/v2/tasks/${pendingAction.taskId}/assignments/${pendingAction.assignmentId}`;

    this.setState({ actionInFlight: true, actionError: undefined });

    const body: { outcome: string; comment?: string } = { outcome: pendingAction.outcome };
    if (pendingAction.comment && pendingAction.comment.trim().length > 0) {
      body.comment = pendingAction.comment.trim();
    }

    try {
      await axios.patch(url, body, {
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${accessToken}`
        }
      });

      this.setState({
        actionInFlight: false,
        pendingAction: undefined,
        actionSuccessMessage: `"${pendingAction.taskName}" completed with outcome: ${pendingAction.outcome}`
      });

      await this.getTasks(this.state.statusFilter);
    }
    catch (error) {
      const message = axios.isAxiosError(error)
        ? (error.response ? JSON.stringify(error.response.data) : error.message)
        : 'An unexpected error occurred.';
      this.setState({
        actionInFlight: false,
        actionError: message
      });
    }
  }

  protected getGreeting(): string {
    const hours = new Date().getHours();
    if (hours < 12) return 'Good morning';
    if (hours <= 18) return 'Good afternoon';
    return 'Good evening';
  }

  public render(): React.ReactElement<INacTaskActionsProps> {
    const { tasks, statusFilter, pendingAction, actionInFlight, actionError, actionSuccessMessage } = this.state;

    return (
      <section className={`${styles.nacTasks} ${this.props.hasTeamsContext ? styles.teams : ''}`}>
        {!this.props.hideHeader && (
          <div className={styles.welcome}>
            <img alt="" src={this.props.isDarkTheme ? nintexReversedLogo : nintexColorLogo} className={styles.welcomeImage} />
            <h2>{escape(this.getGreeting())}, {escape(this.props.userDisplayName)}!</h2>
          </div>
        )}
        <div className='nintex-login'>

          {!this.state.accessToken && !tasks &&
            <button type='button' onClick={this.loginAndGetTasksClick} className={styles.button}>Log In and Get Tasks</button>}

          {actionSuccessMessage && (
            <div className={styles.successBanner}>
              ✅ {actionSuccessMessage}
              <button type="button" className={styles.dismissButton} onClick={() => this.setState({ actionSuccessMessage: undefined })}>✕</button>
            </div>
          )}

          {tasks &&
            <div>
              <h2>Tasks</h2>

              <div className={styles.filterBar}>
                {STATUS_OPTIONS.map((option) => (
                  <button
                    key={option}
                    type="button"
                    className={`${styles.filterButton} ${statusFilter === option ? styles.filterButtonActive : ''}`}
                    onClick={() => this.handleFilterChange(option)}
                  >
                    {option}
                  </button>
                ))}
              </div>

              {tasks.length === 0 && (
                <div className={styles.emptyState}>No tasks found for this filter.</div>
              )}

              <ul className={styles.taskList}>
                {tasks.map((task: ITaskProps) => (
                  <li className={styles.taskCard} key={task.id}>
                    <div className={styles.taskHeader}>
                      <p className={styles.taskTitle}>{task.name}</p>
                      <span className={`${styles.statusBadge} ${getStatusClass(task.status || '')}`}>
                        {task.status}
                      </span>
                    </div>

                    {task.description && (
                      <p className={styles.taskDescription}>{task.description}</p>
                    )}

                    {(() => {
                      const due = task.dueDate ? new Date(task.dueDate) : undefined;
                      const valid = due && !isNaN(due.getTime());
                      return valid
                        ? <p className={styles.taskDueDate}>Due {due.toLocaleDateString()}</p>
                        : <p className={styles.taskDueDate}>No due date</p>;
                    })()}

                    {task.message && (
                      <div
                        className={styles.taskMessage}
                        dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(task.message) }}
                      />
                    )}

                    {task.taskAssignments && task.taskAssignments.map((assignment: ITaskAssignment) => (
                      <div key={assignment.id}>
                        {assignment.urls && assignment.urls.formUrl && (
                          <p className={styles.formLink}>
                            <a href={assignment.urls.formUrl} target="_blank" rel="noreferrer">Open form to respond</a>
                          </p>
                        )}
                      </div>
                    ))}

                    {task.expressApproval === 'include' && (task.validOutcomes || task.outcomes) && ((task.validOutcomes || task.outcomes) || []).length > 0 && (task.status || '').toLowerCase() === 'active' && (
                      <div className={styles.outcomeBar}>
                        {(task.validOutcomes || task.outcomes || []).map((outcome) => (
                          <button
                            key={outcome}
                            type="button"
                            className={`${styles.outcomeButton} ${outcomeButtonStyle(outcome)}`}
                            onClick={() => this.handleOutcomeClick(
                              task.id,
                              task.name,
                              task.taskAssignments && task.taskAssignments.length > 0 ? task.taskAssignments[0].id : undefined,
                              outcome
                            )}
                          >
                            {outcome}
                          </button>
                        ))}
                      </div>
                    )}
                  </li>
                ))}
              </ul>
            </div>
          }
        </div>

        {pendingAction && (
          <div className={styles.confirmOverlay}>
            <div className={styles.confirmDialog}>
              <h3>Confirm action</h3>
              <p>
                This can&apos;t be undone. Are you sure you want to set
                {' '}<strong>&quot;{pendingAction.taskName}&quot;</strong> to outcome
                {' '}<strong>{pendingAction.outcome}</strong>?
              </p>

              {actionError && (
                <div className={styles.errorBanner}>⚠ {actionError}</div>
              )}

              <label className={styles.commentLabel} htmlFor="task-outcome-comment">
                Comment <span className={styles.optionalTag}>(optional)</span>
              </label>
              <textarea
                id="task-outcome-comment"
                className={styles.commentInput}
                placeholder="Add a comment…"
                rows={3}
                value={pendingAction.comment}
                disabled={actionInFlight}
                onChange={(e) => this.handleCommentChange(e.target.value)}
              />

              <div className={styles.confirmActions}>
                <button
                  type="button"
                  className={styles.confirmButton}
                  disabled={actionInFlight}
                  onClick={() => { this.handleConfirmAction().catch((error) => { console.error(error); }); }}
                >
                  {actionInFlight ? 'Submitting...' : `Confirm: ${pendingAction.outcome}`}
                </button>
                <button
                  type="button"
                  className={styles.cancelButton}
                  disabled={actionInFlight}
                  onClick={this.handleCancelAction}
                >
                  Cancel
                </button>
              </div>
            </div>
          </div>
        )}
      </section>
    );
  }
}

function outcomeButtonStyle(outcome: string): string {
  const dict = styles as unknown as Record<string, string>;
  const lower = outcome.toLowerCase();
  if (lower.indexOf('approv') !== -1) return dict.outcomeApprove;
  if (lower.indexOf('reject') !== -1 || lower.indexOf('declin') !== -1) return dict.outcomeReject;
  return dict.outcomeNeutral;
}
