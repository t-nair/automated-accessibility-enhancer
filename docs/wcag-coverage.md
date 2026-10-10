# WCAG 2.1 coverage

Which WCAG 2.1 Level A and AA success criteria the tool covers for slides (`.pptx`) and HTML pages.

- **Fixed**: the tool changes the file so the problem is gone.
- **Reported only**: the tool finds the problem and lists it in the report, but a person has to fix it.
- **Not covered**: not checked yet. The issue column links the planned work, where there is one.

All 50 WCAG 2.1 Level A and AA criteria are listed. HTML support is itself planned ([#41](https://github.com/t-nair/automated-accessibility-enhancer/issues/41)), so every HTML check is "Not covered" today.

Update this table in the same pull request that adds or changes a check.

| Criterion | Name | Level | Status | Notes | Issues |
|---|---|---|---|---|---|
| [1.1.1](https://www.w3.org/WAI/WCAG21/Understanding/non-text-content.html) | Non-text Content | A | Fixed | Slides: pictures without usable alt text get a generated description. Charts, SmartArt, shapes, grouped pictures, decorative marking and master/layout images are not handled. HTML: not covered. | [#29](https://github.com/t-nair/automated-accessibility-enhancer/issues/29), [#30](https://github.com/t-nair/automated-accessibility-enhancer/issues/30), [#31](https://github.com/t-nair/automated-accessibility-enhancer/issues/31), [#32](https://github.com/t-nair/automated-accessibility-enhancer/issues/32), [#33](https://github.com/t-nair/automated-accessibility-enhancer/issues/33), [#42](https://github.com/t-nair/automated-accessibility-enhancer/issues/42) |
| [1.2.1](https://www.w3.org/WAI/WCAG21/Understanding/audio-only-and-video-only-prerecorded.html) | Audio-only and Video-only (Prerecorded) | A | Not covered |  | [#34](https://github.com/t-nair/automated-accessibility-enhancer/issues/34) |
| [1.2.2](https://www.w3.org/WAI/WCAG21/Understanding/captions-prerecorded.html) | Captions (Prerecorded) | A | Not covered |  | [#34](https://github.com/t-nair/automated-accessibility-enhancer/issues/34) |
| [1.2.3](https://www.w3.org/WAI/WCAG21/Understanding/audio-description-or-media-alternative-prerecorded.html) | Audio Description or Media Alternative (Prerecorded) | A | Not covered |  | [#34](https://github.com/t-nair/automated-accessibility-enhancer/issues/34) |
| [1.2.4](https://www.w3.org/WAI/WCAG21/Understanding/captions-live.html) | Captions (Live) | AA | Not covered | Only relevant to live streams embedded in a page. | none yet |
| [1.2.5](https://www.w3.org/WAI/WCAG21/Understanding/audio-description-prerecorded.html) | Audio Description (Prerecorded) | AA | Not covered |  | [#34](https://github.com/t-nair/automated-accessibility-enhancer/issues/34) |
| [1.3.1](https://www.w3.org/WAI/WCAG21/Understanding/info-and-relationships.html) | Info and Relationships | A | Not covered | Table headers, form labels and heading structure are not checked. | [#27](https://github.com/t-nair/automated-accessibility-enhancer/issues/27), [#43](https://github.com/t-nair/automated-accessibility-enhancer/issues/43), [#46](https://github.com/t-nair/automated-accessibility-enhancer/issues/46) |
| [1.3.2](https://www.w3.org/WAI/WCAG21/Understanding/meaningful-sequence.html) | Meaningful Sequence | A | Fixed | Slide titles are moved to the front of the reading order. Other shapes are left as they are. | [#8](https://github.com/t-nair/automated-accessibility-enhancer/issues/8) |
| [1.3.3](https://www.w3.org/WAI/WCAG21/Understanding/sensory-characteristics.html) | Sensory Characteristics | A | Not covered |  | none yet |
| [1.3.4](https://www.w3.org/WAI/WCAG21/Understanding/orientation.html) | Orientation | AA | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| [1.3.5](https://www.w3.org/WAI/WCAG21/Understanding/identify-input-purpose.html) | Identify Input Purpose | AA | Not covered |  | none yet |
| [1.4.1](https://www.w3.org/WAI/WCAG21/Understanding/use-of-color.html) | Use of Color | A | Not covered |  | none yet |
| [1.4.2](https://www.w3.org/WAI/WCAG21/Understanding/audio-control.html) | Audio Control | A | Not covered |  | [#34](https://github.com/t-nair/automated-accessibility-enhancer/issues/34) |
| [1.4.3](https://www.w3.org/WAI/WCAG21/Understanding/contrast-minimum.html) | Contrast (Minimum) | AA | Not covered |  | [#25](https://github.com/t-nair/automated-accessibility-enhancer/issues/25), [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| [1.4.4](https://www.w3.org/WAI/WCAG21/Understanding/resize-text.html) | Resize Text | AA | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| [1.4.5](https://www.w3.org/WAI/WCAG21/Understanding/images-of-text.html) | Images of Text | AA | Not covered |  | none yet |
| [1.4.10](https://www.w3.org/WAI/WCAG21/Understanding/reflow.html) | Reflow | AA | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| [1.4.11](https://www.w3.org/WAI/WCAG21/Understanding/non-text-contrast.html) | Non-text Contrast | AA | Not covered |  | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| [1.4.12](https://www.w3.org/WAI/WCAG21/Understanding/text-spacing.html) | Text Spacing | AA | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| [1.4.13](https://www.w3.org/WAI/WCAG21/Understanding/content-on-hover-or-focus.html) | Content on Hover or Focus | AA | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| [2.1.1](https://www.w3.org/WAI/WCAG21/Understanding/keyboard.html) | Keyboard | A | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| [2.1.2](https://www.w3.org/WAI/WCAG21/Understanding/no-keyboard-trap.html) | No Keyboard Trap | A | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| [2.1.4](https://www.w3.org/WAI/WCAG21/Understanding/character-key-shortcuts.html) | Character Key Shortcuts | A | Not covered |  | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| [2.2.1](https://www.w3.org/WAI/WCAG21/Understanding/timing-adjustable.html) | Timing Adjustable | A | Not covered | Auto-advancing transitions. | [#35](https://github.com/t-nair/automated-accessibility-enhancer/issues/35) |
| [2.2.2](https://www.w3.org/WAI/WCAG21/Understanding/pause-stop-hide.html) | Pause, Stop, Hide | A | Not covered | Animations. | [#35](https://github.com/t-nair/automated-accessibility-enhancer/issues/35) |
| [2.3.1](https://www.w3.org/WAI/WCAG21/Understanding/three-flashes-or-below-threshold.html) | Three Flashes or Below Threshold | A | Not covered |  | none yet |
| [2.4.1](https://www.w3.org/WAI/WCAG21/Understanding/bypass-blocks.html) | Bypass Blocks | A | Not covered | Skip links need a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| [2.4.2](https://www.w3.org/WAI/WCAG21/Understanding/page-titled.html) | Page Titled | A | Reported only | Slides with no title are listed in the report. Titles typed into text boxes are detected. HTML: not covered. | [#7](https://github.com/t-nair/automated-accessibility-enhancer/issues/7), [#23](https://github.com/t-nair/automated-accessibility-enhancer/issues/23) |
| [2.4.3](https://www.w3.org/WAI/WCAG21/Understanding/focus-order.html) | Focus Order | A | Not covered | Slide reading order is covered under 1.3.2. HTML focus order needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| [2.4.4](https://www.w3.org/WAI/WCAG21/Understanding/link-purpose-in-context.html) | Link Purpose (In Context) | A | Not covered |  | [#28](https://github.com/t-nair/automated-accessibility-enhancer/issues/28), [#44](https://github.com/t-nair/automated-accessibility-enhancer/issues/44) |
| [2.4.5](https://www.w3.org/WAI/WCAG21/Understanding/multiple-ways.html) | Multiple Ways | AA | Not covered | Applies to sets of web pages. | none yet |
| [2.4.6](https://www.w3.org/WAI/WCAG21/Understanding/headings-and-labels.html) | Headings and Labels | AA | Not covered |  | [#43](https://github.com/t-nair/automated-accessibility-enhancer/issues/43), [#46](https://github.com/t-nair/automated-accessibility-enhancer/issues/46) |
| [2.4.7](https://www.w3.org/WAI/WCAG21/Understanding/focus-visible.html) | Focus Visible | AA | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| [2.5.1](https://www.w3.org/WAI/WCAG21/Understanding/pointer-gestures.html) | Pointer Gestures | A | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| [2.5.2](https://www.w3.org/WAI/WCAG21/Understanding/pointer-cancellation.html) | Pointer Cancellation | A | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| [2.5.3](https://www.w3.org/WAI/WCAG21/Understanding/label-in-name.html) | Label in Name | A | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| [2.5.4](https://www.w3.org/WAI/WCAG21/Understanding/motion-actuation.html) | Motion Actuation | A | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
| [3.1.1](https://www.w3.org/WAI/WCAG21/Understanding/language-of-page.html) | Language of Page | A | Not covered |  | [#24](https://github.com/t-nair/automated-accessibility-enhancer/issues/24), [#45](https://github.com/t-nair/automated-accessibility-enhancer/issues/45) |
| [3.1.2](https://www.w3.org/WAI/WCAG21/Understanding/language-of-parts.html) | Language of Parts | AA | Not covered |  | none yet |
| [3.2.1](https://www.w3.org/WAI/WCAG21/Understanding/on-focus.html) | On Focus | A | Not covered |  | none yet |
| [3.2.2](https://www.w3.org/WAI/WCAG21/Understanding/on-input.html) | On Input | A | Not covered |  | none yet |
| [3.2.3](https://www.w3.org/WAI/WCAG21/Understanding/consistent-navigation.html) | Consistent Navigation | AA | Not covered | Applies to sets of web pages. | none yet |
| [3.2.4](https://www.w3.org/WAI/WCAG21/Understanding/consistent-identification.html) | Consistent Identification | AA | Not covered | Applies to sets of web pages. | none yet |
| [3.3.1](https://www.w3.org/WAI/WCAG21/Understanding/error-identification.html) | Error Identification | A | Not covered |  | none yet |
| [3.3.2](https://www.w3.org/WAI/WCAG21/Understanding/labels-or-instructions.html) | Labels or Instructions | A | Not covered |  | [#46](https://github.com/t-nair/automated-accessibility-enhancer/issues/46) |
| [3.3.3](https://www.w3.org/WAI/WCAG21/Understanding/error-suggestion.html) | Error Suggestion | AA | Not covered |  | none yet |
| [3.3.4](https://www.w3.org/WAI/WCAG21/Understanding/error-prevention-legal-financial-data.html) | Error Prevention (Legal, Financial, Data) | AA | Not covered |  | none yet |
| [4.1.1](https://www.w3.org/WAI/WCAG21/Understanding/parsing.html) | Parsing | A | Not covered |  | none yet |
| [4.1.2](https://www.w3.org/WAI/WCAG21/Understanding/name-role-value.html) | Name, Role, Value | A | Not covered |  | none yet |
| [4.1.3](https://www.w3.org/WAI/WCAG21/Understanding/status-messages.html) | Status Messages | AA | Not covered | Needs a rendered page. | [#47](https://github.com/t-nair/automated-accessibility-enhancer/issues/47) |
