import React from "react";
import * as du from "../dateUtils";
import { currentLoginHintKey, IMyOwnState, ITeamsAuthComponentState, TeamsBaseComponentWithAuth } from "../msteams-react-base-component-with-auth";
import { AppConfig } from '../../../config/AppConfig';
import { Link } from "../ui";
import { Button, Dialog, DialogActions, DialogContent, DialogContentText, TextField } from "@material-ui/core";
import * as log from '../logger'

/** about(index.html)用ロケール依存リソース定義 */
export interface IFindMsgIndexTranslation {
    /** マイクロソフトアカウントのユーザを入力してください */
    UserPrincipalNameTitle: string;
}

/** about(index.html)用クラス固有ステータス */
interface IFindMsgIndexState extends IMyOwnState {
    /** クリックされたリンクの対象タブ名 */
    selectedTab: string;
    /** ダイアログから入力されたマイクロソフトアカウント名 */
    dialogInput: string;
}


export class FindMsgIndex extends TeamsBaseComponentWithAuth {
    protected showInformation = false;
    protected async setAdditionalState(newstate: ITeamsAuthComponentState, context?: microsoftTeams.Context, inTeams?: boolean): Promise<void> {
        if (context && inTeams? true : false) {
            log.info(`★★★ setAdditionalState is called from componentDidMount; hosted in teams ★★★`);
        }
        const {hostedInTeams, teamsInfo} = newstate;
        let {loginHint} = teamsInfo;
        log.info(`★★★ Initial loginHint: [${loginHint}] ★★★`);
        if (!hostedInTeams && loginHint === "") {
            loginHint = sessionStorage.getItem(currentLoginHintKey) ?? "";
            log.info(`★★★ Fetched saved loginHint from sessionStorage: [${loginHint}] ★★★`);
            newstate.teamsInfo.loginHint = loginHint;
        }
    }

    protected requireDatabase = false;
    protected requireMicrosoftLogin = false;
    protected isUsingStorage = false;
    protected isTeamAndChannelComboIncluded = false;
    protected GetPageTitle(): string {
        return AppConfig.AppInfo.appName;
    }
    protected CreateMyState(): IMyOwnState {
        const res: IFindMsgIndexState = {
            initialized: true,
            selectedTab: "",
            dialogInput: "",
        }
        return res as IMyOwnState;
    }
    protected setMyState(): IMyOwnState {
        return {initialized: true};
    }
    protected renderContentTop(): JSX.Element {
        const config = AppConfig;
        const imgsrc = `${config.AppInfo.logo}`;
        return (
            <header className="l-header">
                <div className="logo">
                    <img src={imgsrc} className="logo"/>
                </div>
                <div className="l-title">
                    <h1>Welcome to <em>{config.AppInfo.appName} </em>!</h1>
                </div>
            </header>
        );
    }
    protected renderContent(): JSX.Element {
        const href = (tab: string, hint: string): string => {
            return `https://${location.hostname}/${tab}/index.html?hint=${hint}&l=ja-jp`;
        };

        const handleClickOpen = (tab: string) => {
            const {teamsInfo, me} = this.state;
            let {selectedTab, dialogInput} = me as IFindMsgIndexState;
            selectedTab = tab;
            dialogInput = teamsInfo.loginHint;

            log.info(`★★★ Link to tab [${selectedTab}] clicked ★★★`);
            
            const myOwn: IFindMsgIndexState = {
                initialized: true,
                selectedTab: selectedTab,
                dialogInput: dialogInput,
            };
            this.setState({dialogOpen: true, me: myOwn});
        };
      
        const handleClose = () => {
            const myOwn: IFindMsgIndexState = {
                initialized: true,
                selectedTab: "",
                dialogInput: "",
            };
            
            this.setState({ dialogOpen: false, me: myOwn});
        };

        const handleOk = () => {
            const {me} = this.state;
            const {selectedTab, dialogInput} = me as IFindMsgIndexState;
            

            log.info(`★★★ Dialog input value: [${dialogInput}] ★★★`);

            let address = "";
            if (!((dialogInput?? "") === "")) {
                address = href(selectedTab, dialogInput);

                log.info(`★★★ Try navigating to [${address}] ★★★`);
                location.href = address;
            }
        };

        return (
            <div>
                <Dialog open={this.state.dialogOpen} onClose={handleClose} aria-labelledby="form-dialog-title">
                <DialogContent>
                    <DialogContentText>
                        {this.state.translation.about.UserPrincipalNameTitle}
                    </DialogContentText>
                    <TextField
                        required
                        autoFocus
                        margin="dense"
                        id="name"
                        type="email"
                        fullWidth
                        defaultValue={(this.state.me as IFindMsgIndexState).dialogInput}
                        onChange={e => this.onDialogTextChange(e)}
                    />
                </DialogContent>
                <DialogActions>
                    <Button onClick={handleClose} color="primary">
                    Cancel
                    </Button>
                    <Button onClick={handleOk} color="primary">
                    OK
                    </Button>
                </DialogActions>
                </Dialog>

                <article className="l-article">
                    <p>
                        <Link onClick={() => handleClickOpen("FindMsgSearchTab")} disabled={this.state.loading}>Channel Message</Link>
                    </p>
                    <p>
                        <Link onClick={() => handleClickOpen("FindMsgSearchChat")} disabled={this.state.loading}>Chat Message</Link>
                    </p>
                    <p>
                        <Link onClick={() => handleClickOpen("FindMsgTopicsTab")} disabled={this.state.loading}>Topics List</Link>
                    </p>
                    <p>
                        <Link onClick={() => handleClickOpen("FindMsgSearchSchedule")} disabled={this.state.loading}>Schedule</Link>
                    </p>

                </article>
            </div>
        );
    }
    protected renderContentBottom(): JSX.Element {
        return (<div/>);
    }
    protected setStateCallBack(): void {
        // 実装なし
    }

    protected onFilterChangedCallBack(): void {
        // 実装なし
    }
    protected onSearchUserChangedCallBack(): void {
        // 実装なし
    }
    protected onTeamOrChannelChangedCallBack(): void {
        // 実装なし
    }
    protected onDateRangeChangedCallBack(): void {
        // 実装なし
    }
    protected async startSync(): Promise<void> {
        // 実装なし
    }
    protected async GetLastSync(): Promise<Date> {
        return du.now();
    }

    onDialogTextChange(e: React.ChangeEvent<HTMLTextAreaElement | HTMLInputElement>): void {
        const data = e.target.value?? "";
        log.info(`★★★ Dialog text changed: [${data}] ★★★`);

        const {me} = this.state;
        const {selectedTab} = me as IFindMsgIndexState;
        const myOwn: IFindMsgIndexState = {
            initialized: true,
            selectedTab: selectedTab,
            dialogInput: data,
        };
        this.setState({ me: myOwn });
    }
}
