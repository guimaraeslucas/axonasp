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
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>AxonASP Database Conversion Tool</title>
        <link rel="stylesheet" href="98.css" />
        <link rel="icon" type="image/x-icon" href="favicon.ico" />
        <link
            rel="icon"
            type="image/png"
            sizes="16x16"
            href="favicon-16x16.png" />
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

            #wizardwindow {
                position: absolute;
                z-index: 9;
            }

            #wizardwindowheader:active {
                cursor: move;
                z-index: 10;
            }
        </style>
    </head>

    <body>
        <div class="center">
            <div
                class="window"
                id="wizardwindow"
                style="width: 650px; user-select: none">
                <div class="title-bar" id="wizardwindowheader">
                    <div class="title-bar-text">
                        <img
                            src="favicon-16x16.png"
                            style="
                                vertical-align: sub;
                                margin: -1px 5px -1px 1px;
                            " />AxonASP Database Conversion Tool
                    </div>
                    <div class="title-bar-controls">
                        <button aria-label="Close" disabled></button>
                    </div>
                </div>
                <div class="window-body">
                    <iframe
                        src="wizard.asp"
                        width="100%"
                        height="503"
                        style="border: none"></iframe>
                </div>
            </div>
        </div>
    </body>
    <script>
        // Make the DIV element draggable:
        dragElement(document.getElementById("wizardwindow"));

        function dragElement(elmnt) {
            var pos1 = 0,
                pos2 = 0,
                pos3 = 0,
                pos4 = 0;
            if (document.getElementById(elmnt.id + "header")) {
                document.getElementById(elmnt.id + "header").onmousedown =
                    dragMouseDown;
            } else {
                elmnt.onmousedown = dragMouseDown;
            }

            function dragMouseDown(e) {
                e = e || window.event;
                e.preventDefault();
                pos3 = e.clientX;
                pos4 = e.clientY;
                document.onmouseup = closeDragElement;
                document.onmousemove = elementDrag;
            }

            function elementDrag(e) {
                e = e || window.event;
                e.preventDefault();
                pos1 = pos3 - e.clientX;
                pos2 = pos4 - e.clientY;
                pos3 = e.clientX;
                pos4 = e.clientY;
                elmnt.style.top = elmnt.offsetTop - pos2 + "px";
                elmnt.style.left = elmnt.offsetLeft - pos1 + "px";
            }

            function closeDragElement() {
                document.onmouseup = null;
                document.onmousemove = null;
            }
        }
    </script>
</html>
