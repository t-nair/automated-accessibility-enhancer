# The pipeline as a Lambda container image.
#
# A container rather than a zip because python-pptx and Pillow both ship
# compiled parts, and building them against the right Lambda runtime is far
# easier inside the image AWS publishes than it is to reproduce in a zip.
FROM public.ecr.aws/lambda/python:3.12

# An older .ppt cannot be opened by python-pptx, so the pipeline converts it
# first by running LibreOffice. Only the presentation parts are installed:
# the full suite is several times the size and nothing here opens a document
# or a spreadsheet.
RUN dnf install -y libreoffice-impress libreoffice-core \
    && dnf clean all \
    && rm -rf /var/cache/dnf

# LibreOffice writes a user profile the first time it runs, and /tmp is the only
# writable place in a Lambda container. Without this the conversion fails with a
# read-only filesystem error rather than anything that explains itself.
ENV HOME=/tmp

# Copied and installed on its own so that editing the pipeline does not
# invalidate the layer holding the dependencies
COPY requirements.txt ${LAMBDA_TASK_ROOT}/
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r ${LAMBDA_TASK_ROOT}/requirements.txt

COPY z_reorder.py lambda_handler.py ${LAMBDA_TASK_ROOT}/

# Captions come from Claude on Bedrock, so nothing is downloaded at runtime and
# a cold start is only the container coming up.
CMD ["lambda_handler.lambda_handler"]
