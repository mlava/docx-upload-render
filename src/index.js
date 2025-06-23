import { renderDocx } from "./docxComponent";
import createButtonObserver from "roamjs-components/dom/createButtonObserver";
import getDropUidOffset from "roamjs-components/dom/getDropUidOffset";
import createBlock from "roamjs-components/writes/createBlock";
import updateBlock from "roamjs-components/writes/updateBlock";
import createHTMLObserver from "roamjs-components/dom/createHTMLObserver";
import getUids from "roamjs-components/dom/getUids";

var runners = {
    observers: [],
}
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
    }

}

function isDocxFile(file) {
    return (
        file &&
        (
            file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
            file.name?.toLowerCase().endsWith(".docx")
        )
    );
}