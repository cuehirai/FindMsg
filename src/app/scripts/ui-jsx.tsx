import React from "react";
import { Backdrop, Box, Button, Chip, Dialog, DialogActions, DialogContent, DialogContentText, FormControlLabel, Switch, Tooltip, Typography, withStyles } from "@material-ui/core";
import { highlightNode, collapse, empty } from "./highlight";
import { collapseConsecutiveChar, sanitize, stripHtml } from "./purify";
import { FindMsgChannel, FindMsgTeam } from "./db";
import * as log from './logger';
import { IMessageTranslation } from "./i18n/IMessageTranslation";
import { db, DbLoginResult } from "./db/Database";
import * as du from "./dateUtils";
import { Client } from "@microsoft/microsoft-graph-client";

/**
 * コンテンツコンポーネント用プロパティ
 */
export interface ContentElementProps {
    body: string;
    type: "text" | "html";
    filter: string;
    tooltip?: boolean;
}

/** ChipsArray生成時の引数用プロパティ */
export interface IChipsArrayProp {
    names: (string | null)[];
    nodata: string;
}

/** チャネルIDから遡ってチャネル名、チーム名を取得するマップ用レコード */
export interface IChannelInfo {
    channelId: string;
    teamId: string;
    channelDisplayName: string;
    teamDisplayName: string;
}

/**
 * コンテンツコンポーネント（表示しきれない場合に「...」で省略したり検索にヒットした部分をハイライトしたい要素に使用）
 * @param param0 
 */
export const ContentElement: React.FunctionComponent<ContentElementProps> = ({ type, body, filter, tooltip }: ContentElementProps) => {
    const el = React.useRef<HTMLSpanElement>(null);
    React.useEffect(() => {
        if (!el.current) return;
        empty(el.current);
        const c = document.createElement("span");
        if (type === "text") {
            c.textContent = body;
        } else {
            // there is no more html, but still entities like &nbsp;
            // c.innerHTML = stripHtml(body);
            c.innerHTML = stripHtml(collapseConsecutiveChar(sanitize(body), "_", 3));
        }
        const hasHighlight = highlightNode(c, [filter, ""]);
        if (body.length > 30 && hasHighlight) collapse(c, 20, 6);
        el.current.appendChild(c);
    }, [type, body, filter, tooltip]);
    let res = <span ref={el} />;
    if (tooltip?? false) {
        res = (
            <HtmlTooltip title={
                <React.Fragment>
                    <span dangerouslySetInnerHTML={{__html: body}} />
                </React.Fragment>
            }>
                <span ref={el} />
            </HtmlTooltip>
        );
    }
    return res;
}

/**
 * 名前の配列からChip配列を生成します。
 * 名前配列に要素がない場合はnodataで指定された文字列を返します。
 * @param param0 
 */
export function ChipsArray({ names, nodata }: IChipsArrayProp): JSX.Element {
    const chipData: Array<JSX.Element> = [];
    if (names.length > 0) {
        names.forEach(name => {
            if (name) {
                chipData.push(<Chip label={name} />);
            }
        });
        return (
            <div className="tooltip">
                {chipData}
            </div>
        );
    } else {
        return (
            <div>
                {nodata}
            </div>
        );
    }
}

/**
 * HTMLを表示するTooltipを生成します。
 */
export const HtmlTooltip = withStyles((theme) => ({
    tooltip: {
        backgroundColor: '#f5f5f9',
        color: 'rgba(0, 0, 0, 0.87)',
        maxWidth: 800,
        fontSize: theme.typography.pxToRem(12),
        border: '1px solid #dadde9',
    },
}))(Tooltip)

/** チャネルIDをキーとしてチャネル名やチーム名を取得できるMAPを生成します（UIでもJSXでもないですが・・） */
export const getChannelMap = async (): Promise<Map<string, IChannelInfo>> => {
    const channelMap = new Map<string, IChannelInfo>();
    const channels = await FindMsgChannel.getAll();
    channels.forEach(async rec => {
        const team = await FindMsgTeam.get(rec.teamId);
        if (team) {
            const channelInfo: IChannelInfo = {
                channelId: rec.id,
                teamId: team.id,
                channelDisplayName: rec.displayName,
                teamDisplayName: team.displayName,
            };
            channelMap.set(rec.id, channelInfo);
        }
    })
    return channelMap;
}


export const getInformation = (showInfo?: boolean): {hasInfo: boolean, info: JSX.Element} => {
    let res = {hasInfo: false, info: (<div/>)};
    try {
        if (showInfo?? true) {
            const request = new XMLHttpRequest();
            const downloadUrl = `https://${location.hostname}/information.txt`
            log.info(`requesting for information: url [${downloadUrl}]`);
            request.open("GET", downloadUrl, false);
            request.send();
            if (request.status < 300) {
                const text = request.response;
                log.info(`information found...content size: [${text.length}]`);
                if (text.length > 0) {
                    res = 
                        {hasInfo: true,
                            info: (<div style={{ width: '100%', maxHeight: 100, overflow: "scroll", 
                                    borderWidth: "medium", borderColor: "lightgray", borderStyle: "solid"}}>
                                    <span dangerouslySetInnerHTML = {{__html: text.toString("utf-8")}} />
                            </div>)
                        };
                }
            } else {
                log.warn(`failed to read information: (${request.status}) [${request.statusText}]`)
            }
        }
    } catch (e) {
        log.warn(`failed to read information: [${e}]`);
    }
    return res;
}

/** エクスポート・インポートの確認ダイアログ、進捗状況マスクの制御用インターフェース */
export interface IExportImportArgs {
    /** 最終同期日時を取得するキー（任意） */
    lastSyncedKey?: string;
    /**
     * エクスポート完了時に呼び出されるコールバック
     * ⇒引数のnewStateをsetStateしてください
     */
    exportCallback: (newState: IExportImportState) => Promise<void>;
    /**
     * インポート完了時に呼び出されるコールバック
     * ※lastSyncedKeyを設定している場合はlastSyncedに指定キーで取得した最終同期日時が引数として渡されます。
     * （lastSyncedKeyを設定していない場合はInvalidDate）
     * ⇒引数のnewState(と必要に応じてlastSynced)をsetStateしてください
     */
    importCallback: (newState: IExportImportState, lastSynced: Date) => Promise<void>;
    /**
     * その他、エクスポート・インポートの制御に必要なステートを更新するためのコールバック
     *  ⇒引数のnewStateをsetStateしてください
     */
    otherCallback: (newState: IExportImportState) => void;
    /** アプリのステートからexportImportStateをそのまま渡してください */
    state: IExportImportState;
    /** 言語依存リソース */
    translate: IMessageTranslation;
    /** エクスポート時に「画像をエクスポートするかどうか」を選択できるようにするにはtrueを設定してください */
    exportOptionAvailable: boolean;
}

/** エクスポート・インポートの確認ダイアログ、進捗状況マスクの制御用ステータス */
export interface IExportImportState {
    /** 
     * db.loginの戻り値（初期値：任意）
     * ※DatabaseLoginによりDBにログインすればアプリで意識する必要はありません
     */
    dblogin: DbLoginResult;
    /** 
     * エクスポート確認ダイアログオープンフラグ（初期値：false）
     * ※同期処理の最後のsetStateで、
     * this.setState({ syncing: false, lastSynced, exportImportState:{...this.state.exportImportState, exportDialog: true} });
     * という風に、「true」を設定してください。
     */
    exportDialog: boolean;
    /** エクスポート時のエラーダイアログオープンフラグ（初期値：false）※アプリで意識する必要はありません */
    exportErrorDialog: boolean;
    /** エクスポート時のエラーダイアログのメッセージ（初期値：""）※アプリで意識する必要はありません */
    exportErrorMsg: string;
    /** 
     * インポート確認ダイアログオープンフラグ（初期値：false）
     * ※DatabaseLoginによりDBにログインすればアプリで意識する必要はありません
     */
    importDialog: boolean;
    /** エクスポート中のバックドロップオープンフラグ（初期値：false）※アプリで意識する必要はありません */
    exporting: boolean;
    /** インポート中のバックドロップオープンフラグ（初期値：false）※アプリで意識する必要はありません */
    importing: boolean;
    /** 現在処理中のテーブル（初期値：""）※アプリで意識する必要はありません */
    processingTable: string;
    /** 処理済みのカウント（初期値：0）※アプリで意識する必要はありません */
    doneCount: number;
    /** 全体のカウント（初期値：0）※アプリで意識する必要はありません */
    allCount: number;
    /** 進捗(%)（初期値：0）※アプリで意識する必要はありません */
    currentProgress: number;
    /** 画像をエクスポートするかどうか（初期値：任意）※初期値の設定以外にアプリで意識する必要はありません */
    exportImages: boolean;
}

/**
 * インポート確認を自動化したデータベースログインシーケンス
 * ※db.loginを直接呼び出さずこのメソッドを呼び出し、戻り値のIExportImportStateでsetStateしてください。
 */
export const DatabaseLogin = async(args: {client: Client, userPrincipalName: string, state: IExportImportState, callback: (newState: IExportImportState) => void}): Promise<IExportImportState> => {
    const res: IExportImportState = {...args.state};

    const resultCallback = (result: DbLoginResult): void => {
        const newState: IExportImportState = {...res};
        if (result === "RECOMMEND_IMPORT" || result === "SHOULD_IMPORT") {
            newState.dblogin = result;
            newState.importDialog = true;
            args.callback(newState);
        }
    }
    res.dblogin = await db.login({client: args.client, userPrincipalName: args.userPrincipalName, resultCallback: resultCallback});
    return res;
}

/**
 * エクスポート・インポートの確認ダイアログ及び進捗状況マスクのコンポーネント
 * @param args 制御用インターフェース
 */
export const ExportImportComponents = (args: IExportImportArgs): JSX.Element => {
    const getMessage = (message: string) => {
        const arr = message.split("\n");
        const res: JSX.Element[] = [];
        arr.map(text => {
            res.push(<p>{text}</p>);
        })
        return res;
    };

    const cloneState = (): IExportImportState => {
        return {...args.state};
    }

    const exportCallback = async (message: number) => {
        const newState = cloneState();
        newState.exporting = false;
        if (message > 0) {
            newState.exportErrorDialog = true;
            newState.exportErrorMsg = args.translate.common.oneDriveQuotaShorts(message);
        }
        await args.exportCallback(newState);
    }

    const importCallback = async () => {
        const newState = cloneState();
        newState.importing = false;

        const lastSynced = args.lastSyncedKey? await db.getLastSync(args.lastSyncedKey) : du.invalidDate();
        await args.importCallback(newState, lastSynced);
    }

    const progressCallback = (tableName: string, done: number, all: number, progress: number) => {
        const newState = cloneState();
        newState.processingTable = tableName;
        newState.doneCount = done;
        newState.allCount = all;
        newState.currentProgress = progress;
        args.otherCallback(newState);
    }

    const handleExportDialogClose = () => {
        const newState = cloneState();
        newState.exportDialog = false;
        args.otherCallback(newState);         
    }

    const handleExportDialogOk = () => {
        args.state.exportDialog = false;
        args.state.exporting = true;
        db.export({includeImages: args.state.exportImages, progressCallback: progressCallback, callback: exportCallback })
        const newState = cloneState();
        args.otherCallback(newState);
    }

    const handleExportErrorDialogClose = () => {
        const newState = cloneState();
        newState.exportErrorDialog = false;
        newState.exportErrorMsg = "";
        args.otherCallback(newState);         
    }

    const importDialogMessage = () => {
        const {dblogin} = args.state;
        let res = ""
        if (dblogin === "RECOMMEND_IMPORT") {
            res = args.translate.common.confirmImportNewerData;
        } else if (dblogin === "SHOULD_IMPORT") {
            res = args.translate.common.confirmImportForNewUser;
        }
        return res;
    }

    const handleImportDialogClose = () => {
        const newState = cloneState();
        newState.importDialog = false;
        args.otherCallback(newState);         
    }

    const handleImportDialogOk = () => {
        args.state.importDialog = false;
        args.state.importing = true;
        db.import({ progressCallback: progressCallback, getLastSync: importCallback });

        const newState = cloneState();
        args.otherCallback(newState);
    }

    const handleImageExportChange = (event: React.ChangeEvent<HTMLInputElement>) => {
        const newState = cloneState();
        newState.exportImages = event.target.checked;
        args.otherCallback(newState)
    }

    const progress = () => {
        let res = "";
        const {exporting, importing, processingTable, doneCount, allCount, currentProgress} = args.state;
        if (processingTable === "") {
            res = args.translate.common.standingBy;
        } else {
            if (exporting) {
                res = args.translate.common.exportProgress(processingTable, doneCount, allCount, currentProgress);
            } else if (importing) {
                res = args.translate.common.importProgress(processingTable, doneCount, allCount, currentProgress);
            }    
        }
        return res;
    }


    const res = (
        <div>
            <Dialog open={args.state.exportDialog} onClose={handleExportDialogClose} aria-labelledby="form-dialog-title">
                <DialogContent>
                    <DialogContentText>
                        {getMessage(args.translate.common.confirmExport)}
                        {args.exportOptionAvailable && getMessage(args.translate.common.confirmExportOption)}
                        {args.exportOptionAvailable && <Box style={{display: 'flex', justifyContent: 'end'}}>
                            <FormControlLabel
                                value="start"
                                control={<Switch color="primary" checked={args.state.exportImages} onChange={handleImageExportChange} />}
                                label={args.translate.common.exportImages}
                                labelPlacement="start"
                            />
                        </Box>}
                    </DialogContentText>
                </DialogContent>
                <DialogActions>
                    <Button onClick={handleExportDialogClose} color="primary">
                        {args.translate.common.no}
                    </Button>
                    <Button onClick={handleExportDialogOk} color="primary">
                        {args.translate.common.yes}
                    </Button>
                </DialogActions>
            </Dialog>

            <Dialog open={args.state.exportErrorDialog} onClose={handleExportErrorDialogClose} aria-labelledby="form-dialog-title">
                <DialogContent>
                    <DialogContentText>
                        {getMessage(args.state.exportErrorMsg)}
                    </DialogContentText>
                </DialogContent>
                <DialogActions>
                    <Button onClick={handleExportErrorDialogClose} color="primary">
                        OK
                    </Button>
                </DialogActions>
            </Dialog>

            <Dialog open={args.state.importDialog} onClose={handleImportDialogClose} aria-labelledby="form-dialog-title">
                <DialogContent>
                    <DialogContentText>
                        {getMessage(importDialogMessage())}
                    </DialogContentText>
                </DialogContent>
                <DialogActions>
                    <Button onClick={handleImportDialogClose} color="primary">
                        {args.translate.common.no}
                    </Button>
                    <Button onClick={handleImportDialogOk} color="primary">
                        {args.translate.common.yes}
                    </Button>
                </DialogActions>
            </Dialog>

            <Backdrop open={args.state.exporting || args.state.importing} style={{zIndex: 100, color: '#fff'}}>
                <Box style={{flexDirection: 'row'}}>
                    <Box display="flex" justifyContent="center">
                        <Typography variant="h4">
                            {/* <CircularProgress color="inherit" /> */}
                            {progress()}
                        </Typography>
                    </Box>
                    <Box display="flex" justifyContent="center">
                        <Typography variant="h4">
                            {args.translate.common.exportImportMessage}
                        </Typography>
                    </Box>
                </Box>
            </Backdrop>
        </div>
    );

    return res;
}