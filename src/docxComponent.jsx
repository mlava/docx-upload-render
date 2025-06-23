import React, { useRef, useEffect, useState } from "react";
import * as docx from "docx-preview";
import { createComponentRender } from "roamjs-components/components/ComponentContainer";

function DocxViewer({ blockUid }) {
    let blockData = window.roamAlphaAPI.data.pull("[:block/string]", `[:block/uid \"${blockUid}\"]`);
    blockData = blockData[":block/string"].split("docxdoc:")[1];
    blockData = blockData.trim();
    var fileUrl = blockData.slice(0, -2);

    const [sections, setSections] = useState([]);
    const [styleNodes, setStyleNodes] = useState([]);
    const [currentPage, setCurrentPage] = useState(0);
    const containerRef = useRef(null);

    // Pagination controls with Home and End
    const PaginationControls = () => (
        <div className="docx-pagination">
            <button onClick={() => setCurrentPage(0)} disabled={currentPage === 0}>Home</button>
            <button onClick={() => setCurrentPage((p) => Math.max(0, p - 1))} disabled={currentPage === 0}>Previous</button>
            <span>Page {currentPage + 1} of {sections.length}</span>
            <button onClick={() => setCurrentPage((p) => Math.min(p + 1, sections.length - 1))} disabled={currentPage === sections.length - 1}>Next</button>
            <button onClick={() => setCurrentPage(sections.length - 1)} disabled={currentPage === sections.length - 1}>End</button>
        </div>
    );

    useEffect(() => {
        if (!fileUrl) return;
        fetch(fileUrl)
            .then(res => res.blob())
            .then(blob => {
                const tempDiv = document.createElement("div");
                docx.renderAsync(blob, tempDiv).then(() => {
                    // Extract <style> tags (docx-preview usually injects them at the wrapper level)
                    const styles = Array.from(tempDiv.querySelectorAll("style")).map(node => node.cloneNode(true));
                    setStyleNodes(styles);

                    // Find all <section class="docx">
                    const sectionNodes = Array.from(tempDiv.querySelectorAll("section.docx"));
                    setSections(sectionNodes.map(node => node.cloneNode(true)));
                    setCurrentPage(0);
                });
            });
    }, [fileUrl]);

    useEffect(() => {
        if (containerRef.current && sections.length > 0) {
            containerRef.current.innerHTML = "";
            // Create a wrapper div
            const wrapper = document.createElement("div");
            wrapper.className = "docx-wrapper";
            // Append all style tags
            styleNodes.forEach(styleNode => wrapper.appendChild(styleNode.cloneNode(true)));
            // Append the current section
            wrapper.appendChild(sections[currentPage].cloneNode(true));
            // Add to container
            containerRef.current.appendChild(wrapper);
        }
    }, [sections, currentPage, styleNodes]);

    return (
        <section className="docx-viewer">
            {sections.length > 1 && <PaginationControls />}
            <div ref={containerRef} className="docx-content" />
            {sections.length > 1 && <PaginationControls />}
        </section>
    );
}

export const renderDocx = createComponentRender(
    ({ blockUid }) => (
        <DocxViewer blockUid={blockUid} />
    )
);
