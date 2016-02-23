// Configuration: Obtain Slack web API token at https://api.slack.com/web
const API_TOKEN = PropertiesService.getScriptProperties().getProperty("slack_api_token");
const BASE_URL = "https://slack.com/api/";

if (!API_TOKEN) {
  throw "You should set 'slack_api_token' property from [File] > [Project properties] > [Script properties]";
}

const FOLDER_NAME = "SlackLogs";

/**** Do not edit below unless you know what you are doing ****/

// Columns where informations are assigned
const COL_LOG_TIMESTAMP = 1;
const COL_LOG_USER = 2;
const COL_LOG_TEXT = 3;
const COL_LOG_RAW_JSON = 4;
const COL_MAX = COL_LOG_RAW_JSON;

// Slack offers 10,000 history logs for free plan teams
const MAX_HISTORY_PAGINATION = 10;
const HISTORY_COUNT_PER_PAGE = 1000;

interface SlackResponse {
  ok: boolean;
  error?: string;
}

// https://api.slack.com/methods/groups.list
interface SlackGroupsListResponse extends SlackResponse {
  groups?: SlackChannel[];
  channels?: SlackChannel[];
}

// https://api.slack.com/methods/channels.history
interface SlackChannelsHistoryResponse extends SlackResponse {
  latest?: string;
  oldest?: string;
  has_more: boolean;
  messages: SlackMessage[];
}

// https://api.slack.com/methods/users.list
interface SlackUsersListResponse extends SlackResponse {
  members: SlackUser[];
}

// https://api.slack.com/types/channel
interface SlackChannel {
  id: string;
  name: string;
  created: number;
}

// https://api.slack.com/events/message
interface SlackMessage {
  type: string;
  ts: string;
  user: string;
  text: string;
  username?: string;
}

// https://api.slack.com/types/user
interface SlackUser {
  id: string;
  name: string;
}

// https://api.slack.com/methods/team.info
interface SlackTeamInfoResponse extends SlackResponse {
  team: {
    id: string;
    name: string;
    domain: string;
  };
}

function StoreLogsDelta() {
  let logger = new SlackChannelHistoryLogger();
  logger.run();
}

class SlackChannelHistoryLogger {
  memberNames: { [id: string]: string } = {};
  teamName: string;

  constructor() {
  }

  requestSlackAPI(path: string, params: { [key: string]: any } = {}): SlackResponse {
    let url = `${BASE_URL}${path}?`;
    let qparams = [`token=${encodeURIComponent(API_TOKEN)}`];
    for (let k in params) {
      qparams.push(`${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`);
    }
    url += qparams.join("&");

    Logger.log(`==> GET ${url}`);

    let resp = UrlFetchApp.fetch(url);
    let data = <SlackResponse>JSON.parse(resp.getContentText());
    if (data.error) {
      throw `GET ${path}: ${data.error}`;
    }
    return data;
  }

  run() {
    let usersResp = <SlackUsersListResponse>this.requestSlackAPI("users.list");
    usersResp.members.forEach((member) => {
      this.memberNames[member.id] = member.name;
    });

    let teamInfoResp = <SlackTeamInfoResponse>this.requestSlackAPI("team.info");
    this.teamName = teamInfoResp.team.name;

    let groupsResp = <SlackGroupsListResponse>this.requestSlackAPI("groups.list");
    for (let ch of groupsResp.groups) {
      this.importChannelHistoryDelta(ch, "groups");
    }
    let channelsResp = <SlackGroupsListResponse>this.requestSlackAPI("channels.list");
    for (let ch of channelsResp.channels) {
      this.importChannelHistoryDelta(ch, "channels");
    }
  }

  getLogsFolder(): GoogleAppsScript.Drive.Folder {
    let folder = DriveApp.getRootFolder();
    let path = [FOLDER_NAME, this.teamName];
    path.forEach((name) => {
      let it = folder.getFoldersByName(name);
      if (it.hasNext()) {
        folder = it.next();
      } else {
        folder = folder.createFolder(name);
      }
    });
    return folder;
  }

  getChannelSheet(ch: SlackChannel, readonly: boolean = false): GoogleAppsScript.Spreadsheet.Sheet {
    let spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
    let sheetByID: { [id: string]: GoogleAppsScript.Spreadsheet.Sheet } = {};

    let spreadsheetName = ch.name;
    let folder = this.getLogsFolder();
    let it = folder.getFilesByName(spreadsheetName);
    if (it.hasNext()) {
      let file = it.next();
      spreadsheet = SpreadsheetApp.openById(file.getId());
    } else {
      if (readonly) return null;

      spreadsheet = SpreadsheetApp.create(spreadsheetName);
      folder.addFile(DriveApp.getFileById(spreadsheet.getId()));
    }

    let sheets = spreadsheet.getSheets();
    sheets.forEach((s: GoogleAppsScript.Spreadsheet.Sheet) => {
      let name = s.getName();
      let m = /^(.+) \((.+)\)$/.exec(name); // eg. "general (C123456)"
      if (!m) return;
      sheetByID[m[2]] = s;
    });

    let sheet = sheetByID[ch.id];
    if (!sheet) {
      if (readonly) return null;
      sheet = spreadsheet.insertSheet();
    }

    let sheetName = `${ch.name} (${ch.id})`;
    if (sheet.getName() !== sheetName) {
      sheet.setName(sheetName);
    }

    return sheet;
  }

  importChannelHistoryDelta(ch: SlackChannel, api: string) {
    Logger.log(`importChannelHistoryDelta ${ch.name} (${ch.id})`);

    let now = new Date();
    let oldest = "1"; // oldest=0 does not work

    let existingSheet = this.getChannelSheet(ch, true);
    if (existingSheet) {
      let lastRow = existingSheet.getLastRow();
      try {
        let data = <SlackMessage>JSON.parse(<string>existingSheet.getRange(lastRow, COL_LOG_RAW_JSON).getValue());
        oldest = data.ts;
      } catch (e) {
        Logger.log(`while trying to parse the latest history item from existing sheet: ${e}`);
      }
    }

    let messages = this.loadMessagesBulk(ch, { oldest: oldest }, api);
    let dateStringToMessages: { [dateString: string]: SlackMessage[] } = {};

    messages.forEach((msg) => {
      let date = new Date(+msg.ts * 1000);
      let dateString = this.formatDate(date);
      if (!dateStringToMessages[dateString]) {
        dateStringToMessages[dateString] = [];
      }
      dateStringToMessages[dateString].push(msg);
    });

    for (let dateString in dateStringToMessages) {
      let sheet = this.getChannelSheet(ch);

      let timezone = sheet.getParent().getSpreadsheetTimeZone();
      let lastTS: number = 0;
      let lastRow = sheet.getLastRow();
      if (lastRow > 0) {
        try {
          let data = <SlackMessage>JSON.parse(<string>sheet.getRange(lastRow, COL_LOG_RAW_JSON).getValue());
          lastTS = +data.ts || 0;
        } catch (_) {
        }
      }

      let rows = dateStringToMessages[dateString].filter((msg) => {
        return +msg.ts > lastTS;
      }).map((msg) => {
        let date = new Date(+msg.ts * 1000);
        return [
          Utilities.formatDate(date, timezone, "yyyy-MM-dd HH:mm:ss"),
          this.memberNames[msg.user] || msg.username,
          this.unescapeMessageText(msg.text),
          JSON.stringify(msg)
        ];
      });
      if (rows.length > 0) {
        let range = sheet.insertRowsAfter(lastRow || 1, rows.length)
          .getRange(lastRow + 1, 1, rows.length, COL_MAX);
        range.setValues(rows);
      }
    }
  }

  formatDate(dt: Date): string {
    return Utilities.formatDate(dt, Session.getScriptTimeZone(), "yyyy-MM");
  }

  loadMessagesBulk(ch: SlackChannel, options: { [key: string]: string | number } = {}, api: string): SlackMessage[] {
    let messages: SlackMessage[] = [];

    // channels.history will return the history from the latest to the oldest.
    // If the result's "has_more" is true, the channel has more older history.
    // In this case, use the result's "latest" value to the channel.history API parameters
    // to obtain the older page, and so on.
    options["count"] = HISTORY_COUNT_PER_PAGE;
    options["channel"] = ch.id;
    let loadSince = (oldest?: string) => {
      if (oldest) {
        options["oldest"] = oldest;
      }
      // order: recent-to-older
      let resp = <SlackChannelsHistoryResponse>this.requestSlackAPI(`${api}.history`, options);
      messages = resp.messages.concat(messages);
      return resp;
    };

    let resp = loadSince();
    let page = 1;
    while (resp.has_more && page <= MAX_HISTORY_PAGINATION) {
      resp = loadSince(resp.messages[0].ts);
      page++;
    }

    // oldest-to-recent
    return messages.reverse();
  }

  unescapeMessageText(text?: string): string {
    return (text || "")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, "'")
      .replace(/&amp;/g, "&")
      .replace(/<@(.+?)>/g, ($0, userID) => {
        let name = this.memberNames[userID];
        return name ? `@${name}` : $0;
      });
  }
}
