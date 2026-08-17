import * as React from 'react';
import styles from './NacForms.module.scss';
import type { INacFormsProps } from './INacFormsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { AuthService } from './AuthService';
import axios from 'axios';
import { IFormProps } from './IFormProps';
import { Config } from './Config';
import nintexColorLogo from '../assets/nintex_RGB_color_500.png';
import nintexReversedLogo from '../assets/nintex_RGB_reversed_500.png';

interface INacFormsState {
  accessToken?: string;
  forms?: IFormProps[];
}

export default class NacForms extends React.Component<INacFormsProps, INacFormsState, {}> {
  private authService: AuthService;
  private authConfig: Config;

  constructor(props: INacFormsProps) {
    super(props);
    this.authConfig = new Config(this.props.clientId, this.props.redirectUri, this.props.region, this.props.useProd);
    this.authService = new AuthService(this.authConfig);
    this.state = {
      accessToken: undefined,
      forms: undefined
    };

    // If the URL contains ?code= this page load is the OAuth callback.
    // Complete the token exchange and fetch forms.
    if (AuthService.isAuthCallback()) {
      this.finishAuthAndGetForms().catch((error) => { console.error(error); });
    }
  }

  // On mount, check localStorage for a valid (non-expired) token from a
  // previous login.  If found, load forms immediately — no "Log In" click
  // needed.  Skipped when this page load is an OAuth callback (?code=),
  // because finishAuthAndGetForms() already handles that path.
  public componentDidMount(): void {
    if (!AuthService.isAuthCallback() && !this.state.accessToken) {
      this.tryExistingSession().catch((error) => {
        console.log('[NacForms] No existing session found');
      });
    }
  }

  public componentDidUpdate(prevProps: INacFormsProps): void {
    if (prevProps.clientId !== this.props.clientId ||
        prevProps.redirectUri !== this.props.redirectUri ||
        prevProps.region !== this.props.region ||
        prevProps.useProd !== this.props.useProd) {
      this.authConfig = new Config(this.props.clientId, this.props.redirectUri, this.props.region, this.props.useProd);
      this.authService = new AuthService(this.authConfig);
    }
  }

  private loginAndGetFormsClick = (): void => {
    this.authService.login().catch((error) => {
      console.error('[NacForms] Login error:', error);
    });
  }

  // Completes the plain SharePoint redirect flow (called when ?code= is in URL).
  //
  // When two webparts share the same page, the first one to call
  // signinCallback() consumes the auth code and state entry from
  // localStorage.  The second one will get "No matching state found in
  // storage".  That is expected — the token is already in localStorage
  // from the first webpart, so we just fall back to reading it.
  private finishAuthAndGetForms = async (): Promise<void> => {
    try {
      await this.authService.handleAuthentication();
    } catch (error) {
      // Another webpart on this page likely already consumed the auth
      // code — the token exchange still happened, so the token should
      // be in localStorage.
      console.log('[NacForms] signinCallback skipped (already handled):', (error as Error).message);
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
        this.getForms().catch((error) => { console.error(error); });
      } else {
        console.error('[NacForms] Access token missing after redirect callback');
      }
    } catch (error) {
      console.error('[NacForms] Error during token retrieval', error);
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
    if (!accessToken) {
      console.log('[NacForms] No valid stored token — trying silent login…');
      accessToken = await this.authService.trySilentLogin();
    }

    // 3. If we got a token (either stored or silently renewed), load data.
    if (accessToken) {
      console.log('[NacForms] Resuming session (no login click needed)');
      this.setState({ accessToken });
      this.getForms().catch((error) => { console.error(error); });
    }
  }

  private getForms = async (): Promise<void> => {
    const region = this.authConfig.getRegion();
    const useProd = this.authConfig.getUseProd();
    const url = `https://${region}.nintex${useProd ? '' : 'test'}.io/workflows/v1/forms`;

    try {
      if (!this.state.accessToken) {
        console.error('[NacForms] Access token missing');
        return;
      }

      const response = await axios.get(url, {
        headers: {
          'Accept': 'application/json',
          'Authorization': `Bearer ${this.state.accessToken}`
        }
      });

      this.setState({ forms: response.data.forms || response.data.items || response.data });
    }
    catch (error) {
      if (axios.isAxiosError(error)) {
        console.error('[NacForms] Get forms error:', error.response ? error.response.data : error.message);
      } else {
        console.error('[NacForms] Get forms error:', error);
      }
    }
  }

  protected getGreeting(): string {
    const hours = new Date().getHours();
    if (hours < 12) return 'Good morning';
    if (hours <= 18) return 'Good afternoon';
    return 'Good evening';
  }

  public render(): React.ReactElement<INacFormsProps> {
    return (
      <section className={`${styles.nacForms} ${this.props.hasTeamsContext ? styles.teams : ''}`}>
        {!this.props.hideHeader && (
          <div className={styles.welcome}>
            <img alt="" src={this.props.isDarkTheme ? nintexReversedLogo : nintexColorLogo} className={styles.welcomeImage} />
            <h2>{escape(this.getGreeting())}, {escape(this.props.userDisplayName)}!</h2>
          </div>
        )}
        <div className='nintex-login'>

          {!this.state.accessToken && !this.state.forms &&
            <button type='button' onClick={this.loginAndGetFormsClick} className={styles.button}>Log In and Get Forms</button>}

          {this.state.forms &&
            <div>
              <h2>Start a New Request</h2>
              <p className={styles.sectionHint}>These forms start a new workflow. Open one to submit a new request.</p>
              {this.state.forms.length === 0 &&
                <p>No forms are currently available to you.</p>}
              <ul className={styles.formList}>
                {this.state.forms.map((form: IFormProps) => {
                  const formUrl = (form.urls && form.urls.formUrl) || form.url || form.formUrl;
                  const formTitle = form.name || form.formName || 'Untitled form';
                  const workflowName = (form.workflow && form.workflow.name) || form.workflowName;
                  console.log(`Form: ${formTitle}, URL: ${formUrl}`);
                  return (
                    <li className={`${styles.formCard} ${formUrl ? styles.formCardClickable : ''}`} key={form.id}>
                      {formUrl ? (
                        <a className={styles.formCardLink} href={formUrl} target="_blank" rel="noreferrer">
                          <span className={styles.formTitleLink}>{formTitle}</span>
                          {workflowName && workflowName !== formTitle &&
                            <span className={styles.formDescription}>{workflowName}</span>}
                          {form.description &&
                            <span className={styles.formDescription}>{form.description}</span>}
                        </a>
                      ) : (
                        <>
                          <p className={styles.formTitle}>{formTitle}</p>
                          {workflowName && workflowName !== formTitle &&
                            <p className={styles.formDescription}>{workflowName}</p>}
                          {form.description &&
                            <p className={styles.formDescription}>{form.description}</p>}
                        </>
                      )}
                    </li>
                  );
                })}
              </ul>
            </div>
          }
        </div>
      </section>
    );
  }
}
