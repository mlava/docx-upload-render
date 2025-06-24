import { renderDocx } from "./docxComponent";
import createButtonObserver from "roamjs-components/dom/createButtonObserver";
import getDropUidOffset from "roamjs-components/dom/getDropUidOffset";
import createBlock from "roamjs-components/writes/createBlock";
import updateBlock from "roamjs-components/writes/updateBlock";
import createHTMLObserver from "roamjs-components/dom/createHTMLObserver";
import getUids from "roamjs-components/dom/getUids";

var runners = {
    observers: [],
};
const dropListeners = [];
const pasteListeners = [];

const clickListener = (e) => {
    const target = e.target;
    if (
        target.tagName === "INPUT" &&
        target.parentElement === document.body &&
        target.type === "file"
    ) {
        target.addEventListener(
            "change",
            (e) => {
                const files = e.target.files;
                if (files && files.length > 0 && isDocxFile(files[0])) {
                    uploadDocx({
                        files,
                        getLoadingUid: () => {
                            const { blockUid } = getUids(textareaRef.current);
                            return updateBlock({
                                text: "Uploading...",
                                uid: blockUid,
                            });
                        },
                        e,
                    });
                }
            },
            true
        );
    }
};

function isDocxFile(file) {
    return (
        file &&
        (
            file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
            file.name?.toLowerCase().endsWith(".docx")
        )
    );
}

export default {
    onload: ({ extensionAPI }) => {
        const onloadArgs = { extensionAPI };
        const docxObserver = createButtonObserver({
            attribute: 'docxdoc',
            render: (b) => {
                renderDocx(b, onloadArgs);
            }
        });
        runners['observers'] = [docxObserver];

        extensionAPI.ui.commandPalette.addCommand({
            label: "Render DOCX files in page",
            callback: () => checkDOCX(null, true, false)
        });
        extensionAPI.ui.commandPalette.addCommand({
            label: "Render DOCX files in focused block",
            callback: () => checkDOCX(null, false, false)
        });
        extensionAPI.ui.commandPalette.addCommand({
            label: "Render DOCX files in focused block and children",
            callback: () => checkDOCX(null, false, true)
        });
        window.roamAlphaAPI.ui.blockContextMenu.addCommand({
            label: "Render DOCX files in block(s)",
            callback: (e) => checkDOCX(e, false, false),
        })
        window.roamAlphaAPI.ui.blockContextMenu.addCommand({
            label: "Render DOCX files in block(s) and children",
            callback: (e) => checkDOCX(e, false, true)
        });
        window.roamAlphaAPI.ui.msContextMenu.addCommand({
            label: "Render DOCX files in selected block(s)",
            callback: (e) => checkDOCX(e, false, false),
        });

        // lots of code borrowed and tweaked from roamjs-components in here, thanks @David @Michael
        const uploadDocx = ({
            files,
            getLoadingUid,
            e,
        }) => {
            if (!files) return;
            const fileToUpload = files[0];
            if (fileToUpload) {
                getLoadingUid().then((uid) => {
                    roamAlphaAPI.file.upload({ file: fileToUpload, toast: { hide: false } })
                        .then((r) => {
                            updateBlock({
                                uid,
                                text: `{{docxdoc: ${r}}}`
                            });
                        })
                        .finally(() => {
                            Array.from(document.getElementsByClassName("dnd-drop-bar"))
                                .map((c) => c)
                                .forEach((c) => (c.style.display = "none"));
                        });
                })
            }
            e.stopPropagation();
            e.stopImmediatePropagation();
            e.preventDefault();
        };

        const textareaRef = {
            current: null,
        };

        createHTMLObserver({
            tag: "DIV",
            className: "dnd-drop-area",
            callback: (d) => {
                const dropHandler = (e) => {
                    const files = e.dataTransfer?.files || null;
                    if (files && files.length > 0 && isDocxFile(files[0])) {
                        uploadDocx({
                            files,
                            getLoadingUid: () => {
                                const { parentUid, offset } = getDropUidOffset(d);
                                return createBlock({
                                    parentUid,
                                    order: offset,
                                    node: { text: "Uploading..." },
                                });
                            },
                            e,
                        });
                    }
                };
                d.addEventListener("drop", dropHandler);
                dropListeners.push({ element: d, handler: dropHandler });
            },
        });

        createHTMLObserver({
            tag: "TEXTAREA",
            className: "rm-block-input",
            callback: (t) => {
                textareaRef.current = t;
                const pasteHandler = (e) => {
                    const files = e.clipboardData?.files || null;
                    if (files && files.length > 0 && isDocxFile(files[0])) {
                        uploadDocx({
                            files,
                            getLoadingUid: () => {
                                const { blockUid } = getUids(t);
                                return updateBlock({
                                    text: "Uploading...",
                                    uid: blockUid,
                                });
                            },
                            e,
                        });
                    }
                };
                t.addEventListener("paste", pasteHandler);
                pasteListeners.push({ element: t, handler: pasteHandler });
            },
        });
        document.body.addEventListener("click", clickListener);

        async function checkDOCX(e, page, children) {
            let uids = await roamAlphaAPI.ui.individualMultiselect.getSelectedUids();
            if (uids.length != 0) { // this is individualMultiselect mode
                for (var i = 0; i < uids.length; i++) {
                    var results = window.roamAlphaAPI.data.pull("[:block/string]", [":block/uid", uids[i]]);
                    await checkStrings(results[":block/string"], uids[i]);
                }
            } else if (e) { // block context or ms context (but not individualMultiselect mode)
                if (e.hasOwnProperty("blocks") && e.blocks.length > 0) { // ms context 
                    for (var i = 0; i < e.blocks.length; i++) {
                        var results = window.roamAlphaAPI.data.pull("[:block/string]", [":block/uid", e.blocks[i]["block-uid"]]);
                        await checkStrings(results[":block/string"], e.blocks[i]["block-uid"]);
                    }
                } else { // block context
                    await checkStrings(e["block-string"], e["block-uid"]);
                    if (children) {
                        var results = window.roamAlphaAPI.data.pull("[:block/string :block/uid {:block/children ...} ]", [":block/uid", e["block-uid"]]);
                        if (results.hasOwnProperty([":block/children"])) {
                            await parseTree(results);
                        }
                    }
                }
            } else { // command palette
                if (page) {
                    let pageUID = await window.roamAlphaAPI.ui.mainWindow.getOpenPageOrBlockUid();
                    let tree = window.roamAlphaAPI.pull(`[ :block/string :block/uid {:block/children ...} ]`, [`:block/uid`, pageUID]);
                    await parseTree(tree);
                } else {
                    let blockUID = window.roamAlphaAPI.ui.getFocusedBlock()["block-uid"];
                    let blocks = window.roamAlphaAPI.data.pull("[:block/string :block/uid {:block/children ...}]", [":block/uid", blockUID]);
                    await checkStrings(blocks[":block/string"], blocks[":block/uid"]);
                    if (children && blocks.hasOwnProperty([":block/children"])) {
                        await parseTree(blocks);
                    }
                }
            }

            async function parseTree(blocks) {
                if (blocks.hasOwnProperty([":block/children"])) {
                    for (var i = 0; i < blocks[":block/children"].length; i++) {
                        await checkStrings(blocks[":block/children"][i][":block/string"], blocks[":block/children"][i][":block/uid"]);
                        if (blocks[":block/children"][i].hasOwnProperty([":block/children"])) {
                            await parseTree(blocks[":block/children"][i]);
                        }
                    }
                }
            }

            async function checkStrings(string, uid) {
                const docxUrlRegex = /https?:\/\/[^\s\)\]\}]+?\.docx(\?[^)\]\s]*)?(#[^)\]\s]*)?/gi;
                const docxMarkdownRegex = /{{docxdoc:\s*([^}]+)}}/gi;

                const markdownUrls = new Set();
                let mdMatch;
                while ((mdMatch = docxMarkdownRegex.exec(string)) !== null) {
                    markdownUrls.add(mdMatch[1].trim());
                }

                let newString = string;
                const alreadyWrapped = new Set(markdownUrls);

                newString = newString.replace(docxUrlRegex, (url) => {
                    if (alreadyWrapped.has(url.trim())) return url;
                    alreadyWrapped.add(url.trim());
                    return `{{docxdoc: ${url}}}`;
                });

                if (newString !== string) {
                    await window.roamAlphaAPI.updateBlock({
                        block: {
                            uid: uid,
                            string: newString
                        }
                    });
                }
            }
            document.querySelector("body")?.click();
        }
    },

    onunload: () => {
        for (let index = 0; index < runners['observers'].length; index++) {
            const element = runners['observers'][index];
            element.disconnect();
        }
        document.body.removeEventListener("click", clickListener);

        dropListeners.forEach(({ element, handler }) => {
            element.removeEventListener("drop", handler);
        });

        pasteListeners.forEach(({ element, handler }) => {
            element.removeEventListener("paste", handler);
        });

        window.roamAlphaAPI.ui.blockContextMenu.removeCommand({
            label: "Render DOCX files in block(s)",
        });
        window.roamAlphaAPI.ui.blockContextMenu.removeCommand({
            label: "Render DOCX files in block(s) and children",
        });
        window.roamAlphaAPI.ui.msContextMenu.removeCommand({
            label: "Render DOCX files in selected block(s)",
        });
    }
};
