/**
 * Antigravity PPTX Studio - Interactive Frontend
 */

document.addEventListener("DOMContentLoaded", () => {
    // DOM Elements
    const generatorForm = document.getElementById("generatorForm");
    const generateBtn = document.getElementById("generateBtn");
    const tabButtons = document.querySelectorAll(".tab-btn");
    const tabPanes = document.querySelectorAll(".tab-pane");
    const dropZone = document.getElementById("dropZone");
    const fileInput = document.getElementById("fileInput");
    const dropContent = document.getElementById("dropContent");
    const fileInfoBox = document.getElementById("fileInfoBox");
    const fileNameDisplay = document.getElementById("fileNameDisplay");
    const fileSizeDisplay = document.getElementById("fileSizeDisplay");
    const removeFileBtn = document.getElementById("removeFileBtn");
    const slideCountSlider = document.getElementById("slideCount");
    const slideCountVal = document.getElementById("slideCountVal");
    const themeCards = document.querySelectorAll(".theme-card");
    const progressBox = document.getElementById("progressBox");
    const progressStatusTitle = document.getElementById("progressStatusTitle");
    const progressStatusSub = document.getElementById("progressStatusSub");
    const resultBox = document.getElementById("resultBox");
    const resultTitle = document.getElementById("resultTitle");
    const resultSubtitle = document.getElementById("resultSubtitle");
    const downloadLink = document.getElementById("downloadLink");
    const resetBtn = document.getElementById("resetBtn");
    const agentCommentaryContainer = document.getElementById("agentCommentaryContainer");
    const agentCommentaryText = document.getElementById("agentCommentaryText");
    const slideCountDisplay = document.getElementById("slideCountDisplay");
    const slidesGrid = document.getElementById("slidesGrid");
    const historyModal = document.getElementById("historyModal");
    const recentHistoryBtn = document.getElementById("recentHistoryBtn");
    const closeHistoryBtn = document.getElementById("closeHistoryBtn");
    const historyList = document.getElementById("historyList");
    const badgeText = document.getElementById("badgeText");
    const badgeDoc = document.getElementById("badgeDoc");
    const badgeUrl = document.getElementById("badgeUrl");
    const textInput = document.getElementById("textInput");
    const urlInput = document.getElementById("urlInput");

    // 1. Tab Switching
    tabButtons.forEach(btn => {
        btn.addEventListener("click", () => {
            tabButtons.forEach(b => b.classList.remove("active"));
            tabPanes.forEach(p => p.classList.remove("active"));

            btn.classList.add("active");
            const targetPane = document.getElementById(btn.dataset.tab);
            if (targetPane) targetPane.classList.add("active");
        });
    });

    // 2. Source Badges & Quick Prompts
    textInput.addEventListener("input", () => {
        if (textInput.value.trim().length > 0) {
            badgeText.textContent = "Provided";
            badgeText.style.background = "#4F46E5";
            badgeText.style.color = "#fff";
        } else {
            badgeText.textContent = "Optional";
            badgeText.style.background = "";
            badgeText.style.color = "";
        }
    });

    urlInput.addEventListener("input", () => {
        if (urlInput.value.trim().length > 0) {
            badgeUrl.textContent = "Provided";
            badgeUrl.style.background = "#4F46E5";
            badgeUrl.style.color = "#fff";
        } else {
            badgeUrl.textContent = "Optional";
            badgeUrl.style.background = "";
            badgeUrl.style.color = "";
        }
    });

    document.querySelectorAll(".sample-tag").forEach(tag => {
        tag.addEventListener("click", () => {
            textInput.value = tag.dataset.sample;
            textInput.dispatchEvent(new Event("input"));
            // Switch to text tab
            document.querySelector('[data-tab="tab-text"]').click();
        });
    });

    // 3. Drag and Drop File Upload
    const formatBytes = (bytes, decimals = 2) => {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const dm = decimals < 0 ? 0 : decimals;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
    };

    const handleFile = (file) => {
        if (!file) return;
        fileNameDisplay.textContent = file.name;
        fileSizeDisplay.textContent = formatBytes(file.size);
        dropContent.classList.add("hidden");
        fileInfoBox.classList.remove("hidden");
        badgeDoc.textContent = "Attached";
        badgeDoc.style.background = "#4F46E5";
        badgeDoc.style.color = "#fff";
    };

    fileInput.addEventListener("change", (e) => {
        if (e.target.files && e.target.files[0]) {
            handleFile(e.target.files[0]);
        }
    });

    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropZone.classList.add("dragover");
    });

    dropZone.addEventListener("dragleave", () => {
        dropZone.classList.remove("dragover");
    });

    dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropZone.classList.remove("dragover");
        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
            fileInput.files = e.dataTransfer.files;
            handleFile(e.dataTransfer.files[0]);
        }
    });

    removeFileBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        fileInput.value = "";
        fileInfoBox.classList.add("hidden");
        dropContent.classList.remove("hidden");
        badgeDoc.textContent = "Optional";
        badgeDoc.style.background = "";
        badgeDoc.style.color = "";
    });

    // 4. Slide Counter Slider
    slideCountSlider.addEventListener("input", (e) => {
        slideCountVal.textContent = `${e.target.value} Slides`;
    });

    // 5. Theme Selection
    themeCards.forEach(card => {
        card.addEventListener("click", () => {
            themeCards.forEach(c => c.classList.remove("selected"));
            card.classList.add("selected");
            const radio = card.querySelector('input[type="radio"]');
            if (radio) radio.checked = true;
        });
    });

    // 6. Stepper State Machine
    const setStepper = (activeStep) => {
        for (let i = 1; i <= 4; i++) {
            const stepEl = document.getElementById(`step${i}`);
            if (!stepEl) continue;
            stepEl.classList.remove("active", "completed");
            if (i < activeStep) {
                stepEl.classList.add("completed");
            } else if (i === activeStep) {
                stepEl.classList.add("active");
            }
        }
    };

    // 7. Form Submission (AJAX)
    generatorForm.addEventListener("submit", async (e) => {
        e.preventDefault();

        // Check if at least one input is provided
        const hasText = textInput.value.trim().length > 0;
        const hasFile = fileInput.files && fileInput.files.length > 0;
        const hasUrl = urlInput.value.trim().length > 0;

        if (!hasText && !hasFile && !hasUrl) {
            alert("Please provide at least one source: Text/Prompt, an uploaded document file, or a URL!");
            return;
        }

        // Show progress box & disable button
        progressBox.classList.remove("hidden");
        resultBox.classList.add("hidden");
        generateBtn.disabled = true;
        setStepper(1);
        progressStatusTitle.textContent = "Ingesting Knowledge Sources...";
        progressStatusSub.textContent = "Parsing documents, crawling web URLs, and consolidating inputs";

        progressBox.scrollIntoView({ behavior: "smooth", block: "nearest" });

        // Simulate stepper transition into Antigravity Agent
        const stepTimer1 = setTimeout(() => {
            setStepper(2);
            progressStatusTitle.textContent = "Spawning Google Antigravity Agent...";
            progressStatusSub.textContent = "Agent is synthesizing context and architecting slide narrative";
        }, 1800);

        const stepTimer2 = setTimeout(() => {
            setStepper(3);
            progressStatusTitle.textContent = "MCP Tool Invocation in Progress...";
            progressStatusSub.textContent = "Antigravity Agent is generating slides via PowerPoint MCP Server";
        }, 4500);

        const formData = new FormData(generatorForm);

        try {
            const response = await fetch("/api/generate", {
                method: "POST",
                body: formData
            });

            clearTimeout(stepTimer1);
            clearTimeout(stepTimer2);

            const data = await response.json();

            if (!response.ok || data.status === "error") {
                throw new Error(data.message || "Failed to generate presentation.");
            }

            // Stepper Step 4: Finished
            setStepper(4);
            progressStatusTitle.textContent = "Presentation Generated!";
            progressStatusSub.textContent = "PowerPoint (.pptx) file is ready for download";

            setTimeout(() => {
                progressBox.classList.add("hidden");
                renderResults(data);
            }, 800);

        } catch (error) {
            clearTimeout(stepTimer1);
            clearTimeout(stepTimer2);
            progressBox.classList.add("hidden");
            alert(`Error during generation: ${error.message}`);
        } finally {
            generateBtn.disabled = false;
        }
    });

    // 8. Render Results Function
    const renderResults = (data) => {
        resultBox.classList.remove("hidden");

        resultTitle.textContent = data.filename.replace(".pptx", "").replace(/_/g, " ");
        resultSubtitle.textContent = `Theme: ${data.theme.replace(/_/g, " ").toUpperCase()} • Total Slides: ${data.total_slides} • Widescreen 16:9`;

        downloadLink.href = data.download_url;
        downloadLink.setAttribute("download", data.filename);

        slideCountDisplay.textContent = data.total_slides || (data.slides_preview ? data.slides_preview.length : 0);

        // Agent Commentary
        if (data.agent_response && data.agent_response.trim()) {
            agentCommentaryContainer.classList.remove("hidden");
            agentCommentaryText.textContent = data.agent_response.trim();
        } else {
            agentCommentaryContainer.classList.add("hidden");
        }

        // Render Slides Grid
        slidesGrid.innerHTML = "";
        const previews = data.slides_preview || [];

        if (previews.length > 0) {
            previews.forEach(s => {
                const card = document.createElement("div");
                card.className = "slide-card";

                const badge = document.createElement("div");
                badge.className = "slide-badge";
                badge.textContent = `SLIDE ${s.slide_number}`;

                const title = document.createElement("div");
                title.className = "slide-title-preview";
                title.textContent = s.title || `Slide ${s.slide_number}`;

                const snippet = document.createElement("div");
                snippet.className = "slide-snippet";
                snippet.textContent = s.content_preview || "Slide contents generated successfully.";

                card.appendChild(badge);
                card.appendChild(title);
                card.appendChild(snippet);

                // Speaker notes accordion if notes exist
                if (s.speaker_notes) {
                    const notesSection = document.createElement("div");
                    notesSection.className = "slide-notes-accordion";

                    const toggleBtn = document.createElement("button");
                    toggleBtn.type = "button";
                    toggleBtn.className = "notes-toggle";
                    toggleBtn.innerHTML = `💬 Speaker Notes ▼`;

                    const notesBody = document.createElement("div");
                    notesBody.className = "notes-body";
                    notesBody.textContent = s.speaker_notes;

                    toggleBtn.addEventListener("click", () => {
                        const isOpen = notesBody.classList.toggle("open");
                        toggleBtn.innerHTML = isOpen ? `💬 Speaker Notes ▲` : `💬 Speaker Notes ▼`;
                    });

                    notesSection.appendChild(toggleBtn);
                    notesSection.appendChild(notesBody);
                    card.appendChild(notesSection);
                }

                slidesGrid.appendChild(card);
            });
        }

        resultBox.scrollIntoView({ behavior: "smooth", block: "start" });
    };

    // 9. Reset Button
    resetBtn.addEventListener("click", () => {
        resultBox.classList.add("hidden");
        generatorForm.scrollIntoView({ behavior: "smooth", block: "start" });
    });

    // 10. Recent History Modal
    recentHistoryBtn.addEventListener("click", async () => {
        historyModal.classList.remove("hidden");
        historyList.innerHTML = `<p class="loading-text">Loading recent decks...</p>`;

        try {
            const resp = await fetch("/api/recent");
            const data = await resp.json();

            if (data.presentations && data.presentations.length > 0) {
                historyList.innerHTML = "";
                data.presentations.forEach(p => {
                    const item = document.createElement("div");
                    item.className = "history-item";
                    item.innerHTML = `
                        <div>
                            <div class="history-name">${p.filename}</div>
                            <div class="history-size">${p.size_kb} KB</div>
                        </div>
                        <a href="${p.download_url}" class="btn-history-dl" download>Download</a>
                    `;
                    historyList.appendChild(item);
                });
            } else {
                historyList.innerHTML = `<p class="loading-text">No presentations generated yet.</p>`;
            }
        } catch (e) {
            historyList.innerHTML = `<p class="loading-text">Failed to load recent presentations.</p>`;
        }
    });

    closeHistoryBtn.addEventListener("click", () => {
        historyModal.classList.add("hidden");
    });

    historyModal.addEventListener("click", (e) => {
        if (e.target === historyModal) {
            historyModal.classList.add("hidden");
        }
    });

    // Check health on load
    fetch("/api/health")
        .then(res => res.json())
        .then(data => {
            const badge = document.getElementById("systemStatus");
            if (badge && data.antigravity_ready) {
                badge.querySelector(".status-label").textContent = "Google Antigravity & MCP Ready";
            }
        })
        .catch(() => {});
});

