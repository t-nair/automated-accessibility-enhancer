const uploadForm = document.getElementById("upload-form");
const dropZone = document.getElementById("drop-zone");
const fileInput = document.getElementById("presentation");
const chosenFile = document.getElementById("chosen-file");
const fileList = document.getElementById("file-list");

const ACCEPTED_TYPES = [".pptx", ".ppt", ".pdf"];

// the files chosen so far, which is what the page shows and what gets submitted.
// Dropping or picking more adds to this rather than replacing it, because someone
// gathering a term's slides a few at a time should not lose the earlier ones.
let chosenFiles = [];


function isAccepted(filename) {
    const lowered = filename.toLowerCase();

    for (let i = 0; i < ACCEPTED_TYPES.length; i++) {
        if (lowered.endsWith(ACCEPTED_TYPES[i])) {
            return true;
        }
    }

    return false;
}


function alreadyChosen(file) {
    for (let i = 0; i < chosenFiles.length; i++) {
        if (chosenFiles[i].name === file.name && chosenFiles[i].size === file.size) {
            return true;
        }
    }

    return false;
}


// the file input cannot be changed a file at a time, so the whole list is rebuilt
// from chosenFiles and handed back to it
function putChosenFilesOnTheInput() {
    const kept = new DataTransfer();

    for (let i = 0; i < chosenFiles.length; i++) {
        kept.items.add(chosenFiles[i]);
    }

    fileInput.files = kept.files;
}


function addFiles(files) {
    for (let i = 0; i < files.length; i++) {
        if (!alreadyChosen(files[i])) {
            chosenFiles.push(files[i]);
        }
    }

    putChosenFilesOnTheInput();
    showChosenFiles();
}


function removeChosenFile(index) {
    chosenFiles.splice(index, 1);
    putChosenFilesOnTheInput();
    showChosenFiles();
}


function formatSize(bytes) {
    const kilobytes = bytes / 1024;

    if (kilobytes < 1024) {
        return Math.round(kilobytes) + " KB";
    }

    return (kilobytes / 1024).toFixed(1) + " MB";
}


function makeFileRow(file, index) {
    const row = document.createElement("li");
    row.className = "file-row";

    const name = document.createElement("span");
    name.className = "file-name";
    name.textContent = file.name;

    const size = document.createElement("span");
    size.className = "file-size";
    size.textContent = formatSize(file.size);

    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "file-remove";
    remove.textContent = "×";
    remove.setAttribute("aria-label", "Remove " + file.name);

    // taking a file back off the list is easy to undo by choosing it again, so it
    // happens on the first press without asking
    remove.addEventListener("click", function () {
        removeChosenFile(index);
    });

    row.appendChild(name);
    row.appendChild(size);
    row.appendChild(remove);

    return row;
}


function showChosenFiles() {
    fileList.textContent = "";

    if (chosenFiles.length === 0) {
        chosenFile.textContent = "No files chosen yet.";
        chosenFile.classList.remove("chosen-file-selected");
        return;
    }

    if (chosenFiles.length === 1) {
        chosenFile.textContent = "1 file chosen";
    } else {
        chosenFile.textContent = chosenFiles.length + " files chosen";
    }

    chosenFile.classList.add("chosen-file-selected");

    for (let i = 0; i < chosenFiles.length; i++) {
        fileList.appendChild(makeFileRow(chosenFiles[i], i));
    }
}


if (dropZone) {
    // the browser opens dropped files in a new tab unless we cancel these two events
    dropZone.addEventListener("dragover", function (event) {
        event.preventDefault();
        dropZone.classList.add("drop-zone-active");
    });

    dropZone.addEventListener("dragleave", function () {
        dropZone.classList.remove("drop-zone-active");
    });

    dropZone.addEventListener("drop", function (event) {
        event.preventDefault();
        dropZone.classList.remove("drop-zone-active");
        addFiles(event.dataTransfer.files);
    });

    // the input holds only what was picked this time, so take those and put the
    // full list back
    fileInput.addEventListener("change", function () {
        addFiles(fileInput.files);
    });
}


const deleteForms = document.querySelectorAll(".delete-form");

// deleting cannot be undone, so check first
for (let i = 0; i < deleteForms.length; i++) {
    deleteForms[i].addEventListener("submit", function (event) {
        const sure = confirm("Delete this submission and its files? This cannot be undone.");

        if (!sure) {
            event.preventDefault();
        }
    });
}


const submissionsTable = document.getElementById("submissions-table");

if (submissionsTable) {
    // refresh the list while files are still waiting or being worked on
    const unfinished = submissionsTable.querySelectorAll(".status-queued, .status-processing");

    if (unfinished.length > 0) {
        setTimeout(function () {
            window.location.reload();
        }, 5000);
    }
}


const progressBar = document.getElementById("progress-bar");

if (progressBar) {
    const progressPercent = document.getElementById("progress-percent");
    // the address comes from the page rather than being built here. Behind the
    // project site the app is served under /app, and a path starting from the
    // site root would ask the landing page for the progress instead.
    const progressUrl = progressBar.getAttribute("data-progress-url");

    function showProgress(done, total) {
        let percent = 0;

        if (total > 0) {
            percent = Math.round((done / total) * 100);
        }

        progressBar.value = percent;
        progressPercent.textContent = percent + "%";
    }

    function checkProgress() {
        fetch(progressUrl)
            .then(function (response) {
                return response.json();
            })
            .then(function (data) {
                const stillWorking = data.status === "processing" || data.status === "queued";

                if (!stillWorking) {
                    // finished (or failed), so reload to show the result
                    window.location.reload();
                    return;
                }

                showProgress(data.done, data.total);
                setTimeout(checkProgress, 1000);
            })
            .catch(function () {
                // if one check fails, just try again shortly
                setTimeout(checkProgress, 2000);
            });
    }

    checkProgress();
}


if (uploadForm) {
    uploadForm.addEventListener("submit", function (event) {
        if (fileInput.files.length === 0) {
            event.preventDefault();
            alert("Please choose at least one file before submitting.");
            return;
        }

        const wrongTypeNames = [];

        for (let i = 0; i < fileInput.files.length; i++) {
            if (!isAccepted(fileInput.files[i].name)) {
                wrongTypeNames.push(fileInput.files[i].name);
            }
        }

        if (wrongTypeNames.length > 0) {
            event.preventDefault();
            alert("Only .pptx, .ppt and .pdf files are accepted. Please remove: " + wrongTypeNames.join(", "));
        }
    });
}


const tabBar = document.getElementById("tab-bar");

// without JavaScript both panels are simply shown one after the other, which reads
// fine. With it, the buttons appear and only one is shown at a time.
if (tabBar) {
    const tabs = tabBar.querySelectorAll(".tab-button");

    function showTab(chosen) {
        for (let i = 0; i < tabs.length; i++) {
            const tab = tabs[i];
            const panel = document.getElementById(tab.getAttribute("aria-controls"));
            const isChosen = (tab === chosen);

            tab.setAttribute("aria-selected", isChosen ? "true" : "false");
            tab.classList.toggle("tab-button-chosen", isChosen);
            tab.tabIndex = isChosen ? 0 : -1;

            if (panel) {
                panel.hidden = !isChosen;
            }
        }
    }

    for (let i = 0; i < tabs.length; i++) {
        tabs[i].addEventListener("click", function () {
            showTab(tabs[i]);
        });
    }

    // left and right arrows move between the tabs, which is how tabs usually behave
    tabBar.addEventListener("keydown", function (event) {
        let current = -1;

        for (let i = 0; i < tabs.length; i++) {
            if (tabs[i].getAttribute("aria-selected") === "true") {
                current = i;
            }
        }

        if (event.key === "ArrowLeft" && current > 0) {
            showTab(tabs[current - 1]);
            tabs[current - 1].focus();
        }

        if (event.key === "ArrowRight" && current < tabs.length - 1) {
            showTab(tabs[current + 1]);
            tabs[current + 1].focus();
        }
    });

    tabBar.hidden = false;
    showTab(tabs[0]);
}


const slideViewer = document.getElementById("slide-viewer");

// without JavaScript every slide is shown one after another, which still works.
// With it, one at a time, with buttons to move between them.
if (slideViewer) {
    const panels = slideViewer.querySelectorAll(".slide-panel");
    const slideNav = document.getElementById("slide-nav");
    const counter = document.getElementById("slide-counter");
    const previousButton = document.getElementById("slide-prev");
    const nextButton = document.getElementById("slide-next");
    // the PDF review steps through pages rather than slides
    const label = slideViewer.getAttribute("data-label") || "Slide";
    let currentSlide = 0;

    function showSlide(index) {
        for (let i = 0; i < panels.length; i++) {
            panels[i].hidden = (i !== index);
        }

        counter.textContent = label + " " + panels[index].getAttribute("data-slide") + " of " + panels.length;
        previousButton.disabled = (index === 0);
        nextButton.disabled = (index === panels.length - 1);
        currentSlide = index;
    }

    if (panels.length > 0) {
        slideNav.hidden = false;
        showSlide(0);

        previousButton.addEventListener("click", function () {
            if (currentSlide > 0) {
                showSlide(currentSlide - 1);
            }
        });

        nextButton.addEventListener("click", function () {
            if (currentSlide < panels.length - 1) {
                showSlide(currentSlide + 1);
            }
        });

        // the left and right arrow keys also move between slides
        slideViewer.addEventListener("keydown", function (event) {
            // arrow keys move the cursor while typing a description, not the page
            if (event.target.closest("input, textarea, select")) {
                return;
            }

            if (event.key === "ArrowLeft" && currentSlide > 0) {
                showSlide(currentSlide - 1);
            }

            if (event.key === "ArrowRight" && currentSlide < panels.length - 1) {
                showSlide(currentSlide + 1);
            }
        });
    }
}


// On the PDF review page, clicking an outlined image shows only that image's
// description next to the page. Without JavaScript every description is shown
// and the outlines are links that jump to them, which still works.
const reviewPanels = document.querySelectorAll(".review-layout");

for (let p = 0; p < reviewPanels.length; p++) {
    const boxes = reviewPanels[p].querySelectorAll(".review-box");
    const editors = reviewPanels[p].querySelectorAll(".review-editor");

    const showEditor = function (editorId) {
        for (let i = 0; i < editors.length; i++) {
            editors[i].hidden = (editors[i].id !== editorId);
        }

        for (let i = 0; i < boxes.length; i++) {
            const isChosen = boxes[i].getAttribute("data-editor") === editorId;
            boxes[i].classList.toggle("review-box-chosen", isChosen);
            boxes[i].setAttribute("aria-current", isChosen ? "true" : "false");
        }
    };

    for (let i = 0; i < boxes.length; i++) {
        boxes[i].addEventListener("click", function (event) {
            event.preventDefault();
            const editorId = boxes[i].getAttribute("data-editor");
            showEditor(editorId);
            document.getElementById(editorId).querySelector("textarea").focus();
        });
    }

    if (boxes.length > 0) {
        showEditor(boxes[0].getAttribute("data-editor"));
    }
}
