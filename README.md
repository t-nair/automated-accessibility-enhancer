# automated-accessibility-enhancer

## About
`automated-accessibility-enhancer` is a Python-based tool that processes `.pptx` files to improve their accessibility by automatically tagging headings and correcting reading order. The script takes a directory of PowerPoint files as input and outputs a new directory of updated, accessibility-enhanced presentations, while preserving original layout and visual design.

This project was developed as part of the **University of Washington Advancing Accessibility for Engineering Education** initiative, under the guidance of **Prof. Anant** and **Dr. Lalitha Subramanian**. The work has been presented at **EDUCAUSE 2025 Online** as part of a broader effort to scale accessibility remediation in instructional materials.

Feedback on the design, architecture, or implementation is very welcome. You can reach me via [email](mailto:tanyan07@uw.edu) or [LinkedIn](https://www.linkedin.com/in/tanya-nair-617473287/).

---

## Workflow
![Workflow overview](https://github.com/user-attachments/assets/46f6381d-969f-4e2f-aa14-19f6b2b116c3)

At a high level, the tool:
1. Ingests a directory of `.pptx` files
2. Parses slide structure and elements
3. Automatically identifies and tags headings
4. Updates reading order to improve screen reader navigation
5. Writes updated `.pptx` files to a separate output directory

The core processing logic is implemented in Python and is designed to be deterministic and non-destructive to slide content and aesthetics.

---

## Possible Improvements
Planned or exploratory enhancements include:
- **Image alt-text generation** using vision-language models
- **Support for additional file types**, such as PDFs or exported Prezi files
- **Configurable rules** for heading detection and structure inference
- **Batch processing safeguards**, including validation and rollback
- **Optional GUI or CLI wrapper** for easier use by non-technical users
- **Integration into automation platforms** (e.g., self-hosted workflow engines)

---

## Current Challenges / Barriers
Some of the main challenges encountered so far include:
- **File I/O and execution context constraints** when integrating Python-based file processing into external automation workflows
- **Handling binary PPTX data reliably** across different orchestration environments
- **Scalability and storage management**, especially for batch uploads
- **Platform limitations** in free or self-hosted automation tools compared to managed cloud offerings
- **Balancing automation with accuracy**, particularly for complex slide layouts

These constraints are actively being explored, and design tradeoffs are documented as the project evolves.

---

## Status
The core Python script is functional for local, directory-based processing. Work is ongoing to improve robustness, extensibility, and integration into larger accessibility pipelines.

Contributions, feedback, and design discussions are encouraged.
