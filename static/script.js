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
