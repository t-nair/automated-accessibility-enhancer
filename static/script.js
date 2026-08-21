const uploadForm = document.getElementById("upload-form");
const dropZone = document.getElementById("drop-zone");
const fileInput = document.getElementById("presentation");
const chosenFile = document.getElementById("chosen-file");


function showChosenFile() {
    if (fileInput.files.length === 0) {
        chosenFile.textContent = "No files chosen yet.";
        chosenFile.classList.remove("chosen-file-selected");
        return;
    }

    const names = [];

    for (let i = 0; i < fileInput.files.length; i++) {
        names.push(fileInput.files[i].name);
    }

    if (names.length === 1) {
        chosenFile.textContent = "Chosen file: " + names[0];
    } else {
        chosenFile.textContent = "Chosen " + names.length + " files: " + names.join(", ");
    }

    chosenFile.classList.add("chosen-file-selected");
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

        if (event.dataTransfer.files.length > 0) {
            fileInput.files = event.dataTransfer.files;
            showChosenFile();
        }
    });

    fileInput.addEventListener("change", showChosenFile);
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
    const submissionId = progressBar.getAttribute("data-submission-id");

    function showProgress(done, total) {
        let percent = 0;

        if (total > 0) {
            percent = Math.round((done / total) * 100);
        }

        progressBar.value = percent;
        progressPercent.textContent = percent + "%";
    }

    function checkProgress() {
        fetch("/progress/" + submissionId)
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
            const fileName = fileInput.files[i].name.toLowerCase();

            if (!fileName.endsWith(".pptx") && !fileName.endsWith(".ppt")) {
                wrongTypeNames.push(fileInput.files[i].name);
            }
        }

        if (wrongTypeNames.length > 0) {
            event.preventDefault();
            alert("Only .pptx and .ppt files are accepted. Please remove: " + wrongTypeNames.join(", "));
        }
    });
}
