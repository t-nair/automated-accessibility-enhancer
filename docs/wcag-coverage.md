# WCAG 2.1 coverage

Which WCAG 2.1 Level A and AA success criteria the tool covers for slides (`.pptx`) and HTML pages.

- **Fixed**: the tool changes the file so the problem is gone.
- **Reported only**: the tool finds the problem and lists it in the report, but a person has to fix it.
- **Not covered**: not checked yet. The issue column links the planned work, where there is one.

Criteria that cannot occur in slides or static HTML (live captions, multi-page navigation, pointer gestures and similar) are left out. HTML support is itself planned ([#41](https://github.com/t-nair/automated-accessibility-enhancer/issues/41)), so every HTML check is "Not covered" today.

Update this table in the same pull request that adds or changes a check.

| Criterion | Name | Level | Status | Notes | Issues |
|---|---|---|---|---|---|
| 1.1.1 | Non-text Content | A | Fixed | Slides: pictures without usable alt text get a generated description. Charts, SmartArt, shapes, grouped pictures, decorative marking and master/layout images are not handled. HTML: not covered. | [#29](https://github.com/t-nair/automated-accessibility-enhancer/issues/29), [#30](https://github.com/t-nair/automated-accessibility-enhancer/issues/30), [#31](https://github.com/t-nair/automated-accessibility-enhancer/issues/31), [#32](https://github.com/t-nair/automated-accessibility-enhancer/issues/32), [#33](https://github.com/t-nair/automated-accessibility-enhancer/issues/33), [#42](https://github.com/t-nair/automated-accessibility-enhancer/issues/42) |
| 1.2.1 | Audio-only and Video-only (Prerecorded) | A | Not covered |  | [#34](https://github.com/t-nair/automated-accessibility-enhancer/issues/34) |
| 1.2.2 | Captions (Prerecorded) | A | Not covered |  | [#34](https://github.com/t-nair/automated-accessibility-enhancer/issues/34) |
| 1.2.3 | Audio Description or Media Alternative (Prerecorded) | A | Not covered |  | [#34](https://github.com/t-nair/automated-accessibility-enhancer/issues/34) |
| 1.2.5 | Audio Description (Prerecorded) | AA | Not covered |  | [#34](https://github.com/t-nair/automated-accessibility-enhancer/issues/34) |
| 1.3.1 | Info and Relationships | A | Not covered | Table headers, form labels and heading structure are not checked. | [#27](https://github.com/t-nair/automated-accessibility-enhancer/issues/27), [#43](https://github.com/t-nair/automated-accessibility-enhancer/issues/43), [#46](https://github.com/t-nair/automated-accessibility-enhancer/issues/46) |
| 1.3.2 | Meaningful Sequence | A | Fixed | Slide titles are moved to the front of the reading order. Other shapes are left as they are. | [#8](https://github.com/t-nair/automated-accessibility-enhancer/issues/8) |
| 1.3.3 | Sensory Characteristics | A | Not covered |  | none yet |
| 1.3.4 | Orientation | AA | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| 1.3.5 | Identify Input Purpose | AA | Not covered |  | none yet |
| 1.4.1 | Use of Color | A | Not covered |  | none yet |
| 1.4.2 | Audio Control | A | Not covered |  | [#34](https://github.com/t-nair/automated-accessibility-enhancer/issues/34) |
| 1.4.3 | Contrast (Minimum) | AA | Not covered |  | [#25](https://github.com/t-nair/automated-accessibility-enhancer/issues/25), [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| 1.4.4 | Resize Text | AA | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| 1.4.5 | Images of Text | AA | Not covered |  | none yet |
| 1.4.10 | Reflow | AA | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| 1.4.11 | Non-text Contrast | AA | Not covered |  | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| 1.4.12 | Text Spacing | AA | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| 1.4.13 | Content on Hover or Focus | AA | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| 2.1.1 | Keyboard | A | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| 2.1.2 | No Keyboard Trap | A | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| 2.1.4 | Character Key Shortcuts | A | Not covered |  | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| 2.2.1 | Timing Adjustable | A | Not covered | Auto-advancing transitions. | [#35](https://github.com/t-nair/automated-accessibility-enhancer/issues/35) |
| 2.2.2 | Pause, Stop, Hide | A | Not covered | Animations. | [#35](https://github.com/t-nair/automated-accessibility-enhancer/issues/35) |
| 2.3.1 | Three Flashes or Below Threshold | A | Not covered |  | none yet |
| 2.4.1 | Bypass Blocks | A | Not covered | Skip links need a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| 2.4.2 | Page Titled | A | Reported only | Slides with no title are listed in the report. Titles typed into text boxes are detected. HTML: not covered. | [#7](https://github.com/t-nair/automated-accessibility-enhancer/issues/7), [#23](https://github.com/t-nair/automated-accessibility-enhancer/issues/23) |
| 2.4.3 | Focus Order | A | Not covered | Slide reading order is covered under 1.3.2. HTML focus order needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| 2.4.4 | Link Purpose (In Context) | A | Not covered |  | [#28](https://github.com/t-nair/automated-accessibility-enhancer/issues/28), [#44](https://github.com/t-nair/automated-accessibility-enhancer/issues/44) |
| 2.4.6 | Headings and Labels | AA | Not covered |  | [#43](https://github.com/t-nair/automated-accessibility-enhancer/issues/43), [#46](https://github.com/t-nair/automated-accessibility-enhancer/issues/46) |
| 2.4.7 | Focus Visible | AA | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| 3.1.1 | Language of Page | A | Not covered |  | [#24](https://github.com/t-nair/automated-accessibility-enhancer/issues/24), [#45](https://github.com/t-nair/automated-accessibility-enhancer/issues/45) |
| 3.1.2 | Language of Parts | AA | Not covered |  | none yet |
| 3.2.1 | On Focus | A | Not covered |  | none yet |
| 3.2.2 | On Input | A | Not covered |  | none yet |
| 3.3.1 | Error Identification | A | Not covered |  | none yet |
| 3.3.2 | Labels or Instructions | A | Not covered |  | [#46](https://github.com/t-nair/automated-accessibility-enhancer/issues/46) |
| 3.3.3 | Error Suggestion | AA | Not covered |  | none yet |
| 3.3.4 | Error Prevention (Legal, Financial, Data) | AA | Not covered |  | none yet |
| 4.1.1 | Parsing | A | Not covered |  | none yet |
| 4.1.2 | Name, Role, Value | A | Not covered |  | none yet |
