import * as express from "express";
let fs = require("fs");
let path = require("path");
import * as config from "config";

export class DefaultTab {

    public static getRequestHandler(): express.RequestHandler {
        return async function (req: any, res: any, next: any): Promise<void> {
            try {
                let htmlPage = `<!DOCTYPE html>
                    <html>
                    <head>
                        <title>Bot Info</title>
                        <meta charset="utf-8" />
                        <meta name="viewport" content="width=device-width, initial-scale=1.0">
                        <script src='https://statics.teams.microsoft.com/sdk/v1.0/js/MicrosoftTeams.min.js'></script>
                        <script src='https://code.jquery.com/jquery-1.11.3.min.js'></script>
                    </head>

                    <body>
                        <p>
                            Welcome to SamePage!
                        </p>`;

                htmlPage += `

                    <br>
                    <p id="currentTheme">Reporting</p>
                    <br>
                    <button onclick="showAllCommands()">Click to give update</button>
                    <br>
                    <p>Dashboard</p>
                    <button onclick="getDeeplink()">Click to get digest of team member's plans, problems and progress</button>
                    <br>
                    <p>Settings</p>
                    <br>
                    <button onclick="showContext()">Click for settings</button>
                    <p id="contextOutput"></p>
                    <script>
                        var microsoftTeams;

                        $(document).ready(function () {
                            microsoftTeams.initialize();
                            microsoftTeams.registerOnThemeChangeHandler(function(theme) {
                                document.getElementById('currentTheme').innerHTML = theme;
                            });
                        });

                        function showAllCommands() {
                            window.location = "${config.get("app.baseUri") + "/allCommands"}";
                        }

                        function getDeeplink() {
                            microsoftTeams.shareDeepLink({subEntityId: 'stuff', subEntityLabel: 'stuff2'});
                        }

                        function showContext() {
                            microsoftTeams.getContext((context) => {
                                document.getElementById('contextOutput').innerHTML = JSON.stringify(context);
                            });
                        }
                    </script>
                    </body>
                    </html>`;

                let files = fs.readdirSync("./public/exampleDialogs");
                for (let i = 0; i < files.length; i++) {
                    let currFile = files[i];
                    let fileName = path.parse(currFile).name;
                    let fileExtension = path.parse(currFile).ext;
                    if (fileExtension === ".txt") {
                        htmlPage += `<br>
                        <a href="/exampleDialogs/` + fileName +  `.txt">` + fileName + `.ts</a>`;
                    }
                }

                res.send(htmlPage);
            } catch (e) {
                // Don't log expected errors - error is probably from there not being example dialogs
                res.send(`<html>
                    <body>
                    <p>
                        Sorry. There are no example dialogs to display.
                    </p>
                    <br>
                    <img src="/tab/error_generic.png" alt="default image" />
                    </body>
                    </html>`);
            }
        };
    }
}
