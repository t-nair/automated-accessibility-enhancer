# The pipeline as a Lambda container image.
#
# Built on Debian rather than the image AWS publishes for Lambda, because an older
# .ppt has to be converted by LibreOffice before python-pptx can open it, and the
# AWS image is Amazon Linux 2023, which does not package LibreOffice at all. The
# first build tried `dnf install libreoffice-impress` there and got "No package
# matches". Debian packages it, so the conversion works with an ordinary install.
#
# Lambda runs images that were not built by AWS as long as they include AWS's
# runtime interface client, which is what lets Lambda hand this container an
# event and read back the result. That is the awslambdaric install below.
FROM python:3.12-slim-bookworm

# Only the presentation part of LibreOffice, and none of what it merely
# recommends, which would pull in the rest of the office suite and a desktop.
RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-impress \
    && rm -rf /var/lib/apt/lists/*

# LibreOffice writes a user profile the first time it runs, and /tmp is the only
# writable place in a Lambda container. Without this the conversion fails with a
# read-only filesystem error rather than anything that explains itself.
ENV HOME=/tmp

# log to the console rather than a file, because the console is the only thing
# AWS collects from a container
ENV LOG_FILE=

WORKDIR /var/task

# Copied and installed on its own so that editing the pipeline does not
# invalidate the layer holding the dependencies
COPY requirements.txt ./
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt awslambdaric

COPY z_reorder.py lambda_handler.py ./

# awslambdaric is the part that talks to Lambda. It is given the handler to call
# for each event, in the same module.function form Lambda uses everywhere else.
ENTRYPOINT ["python", "-m", "awslambdaric"]
CMD ["lambda_handler.lambda_handler"]
