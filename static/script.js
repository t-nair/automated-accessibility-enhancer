const uploadForm = document.getElementById("upload-form");
const dropZone = document.getElementById("drop-zone");
const fileInput = document.getElementById("presentation");
const chosenCount = document.getElementById("chosen-count");
const fileList = document.getElementById("file-list");


function formatSize(bytes) {
    const kilobytes = bytes / 1024;

    if (kilobytes < 1024) {
        return Math.round(kilobytes) + " KB";
    }

    const megabytes = kilobytes / 1024;

    return megabytes.toFixed(1) + " MB";
}


function getFileLabel(filename) {
    if (filename.toLowerCase().endsWith(".ppt")) {
        return "PPT";
    }

    return "PPTX";
}


// take the chosen file out of the list, which means rebuilding it because
// the browser does not let us change the list directly
function removeChosenFile(indexToRemove) {
    const kept = new DataTransfer();

    for (let i = 0; i < fileInput.files.length; i++) {
        if (i !== indexToRemove) {
            kept.items.add(fileInput.files[i]);
        }
    }

    fileInput.files = kept.files;
    showChosenFiles();
}


function makeFileRow(file, index) {
    const row = document.createElement("li");
    row.className = "file-row";

    const icon = document.createElement("span");
    icon.className = "file-icon";
    icon.setAttribute("aria-hidden", "true");
    icon.textContent = getFileLabel(file.name);

    const details = document.createElement("span");
    details.className = "file-details";

    const name = document.createElement("span");
    name.className = "file-name";
    name.textContent = file.name;

    const size = document.createElement("span");
    size.className = "file-size";
    size.textContent = formatSize(file.size);

    details.appendChild(name);
    details.appendChild(size);

    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.className = "file-remove";
    removeButton.textContent = "×";
    removeButton.setAttribute("aria-label", "Remove " + file.name);

    removeButton.addEventListener("click", function () {
        removeChosenFile(index);
    });

    row.appendChild(icon);
    row.appendChild(details);
    row.appendChild(removeButton);

    return row;
}


function showChosenFiles() {
    fileList.textContent = "";

    if (fileInput.files.length === 0) {
        chosenCount.textContent = "No files chosen yet.";
        chosenCount.classList.remove("chosen-count-selected");
        return;
    }

    if (fileInput.files.length === 1) {
        chosenCount.textContent = "1 file chosen";
    } else {
        chosenCount.textContent = fileInput.files.length + " files chosen";
    }

    chosenCount.classList.add("chosen-count-selected");

    for (let i = 0; i < fileInput.files.length; i++) {
        fileList.appendChild(makeFileRow(fileInput.files[i], i));
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

        if (event.dataTransfer.files.length > 0) {
            fileInput.files = event.dataTransfer.files;
            showChosenFiles();
        }
    });

    fileInput.addEventListener("change", showChosenFiles);
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


const slideViewer = document.getElementById("slide-viewer");

// without JavaScript every slide is simply shown one after another, which still works.
// With it, we show one at a time and add buttons to move between them.
if (slideViewer) {
    const panels = slideViewer.querySelectorAll(".slide-panel");
    const slideNav = document.getElementById("slide-nav");
    const counter = document.getElementById("slide-counter");
    const previousButton = document.getElementById("slide-prev");
    const nextButton = document.getElementById("slide-next");
    let currentSlide = 0;

    function showSlide(index) {
        for (let i = 0; i < panels.length; i++) {
            panels[i].hidden = (i !== index);
        }

        const panel = panels[index];
        counter.textContent = "Slide " + panel.getAttribute("data-slide") +
            " of " + panels.length + " - " + panel.getAttribute("data-note");

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
            if (event.key === "ArrowLeft" && currentSlide > 0) {
                showSlide(currentSlide - 1);
            }

            if (event.key === "ArrowRight" && currentSlide < panels.length - 1) {
                showSlide(currentSlide + 1);
            }
        });
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
