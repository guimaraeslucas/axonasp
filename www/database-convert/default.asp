<html lang="en">
    <!--
        
        AxonASP Server
        Copyright (C) 2026 G3pix Ltda. All rights reserved.
        
        Developed by Lucas Guimarães - G3pix Ltda
        Contact: https://g3pix.com.br/
        Project URL: https://g3pix.com.br/axonasp
        
        This Source Code Form is subject to the terms of the Mozilla Public
        License, v. 2.0. If a copy of the MPL was not distributed with this
        file, You can obtain one at https://mozilla.org/MPL/2.0/.
        
        Attribution Notice:
        If this software is used in other projects, the name "AxonASP Server"
        must be cited in the documentation or "About" section.
        
        Contribution Policy:
        Modifications to the core source code of AxonASP Server must be
        made available under this same license terms.
        
        -->
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>AxonASP Database Conversion Tool</title>
        <link rel="stylesheet" href="98.css">
        <style>
            body,
            html {
                background-color: #396da7;
                margin: 0;
                font-size: 12px;
                font-family: Tahoma, sans-serif;
                font-size: 9pt;
                height: 100%;
                background-position: 97% 3%;
                background-repeat: no-repeat;
                background-size: 280px;
                background-blend-mode: hard-light;
                background-image: url(axon.png);
            }

            .center {
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100%;
            }
        </style>
    </head>

    <body>
        <div class="center">
            <div class="window" style="width: 650px; user-select: none;">
                <div class="title-bar">
                    <div class="title-bar-text"><img src="world.png"
                            style="vertical-align: sub;margin: -1px 5px -1px 1px;">AxonASP Database Conversion Tool
                    </div>
                    <div class="title-bar-controls">
                        <button aria-label="Close" disabled></button>
                    </div>
                </div>
                <div class="window-body">
                    <iframe src="wizard.asp" width="100%" height="503" style="border:none;"></iframe>
                </div>
            </div>
        </div>
    </body>

</html>