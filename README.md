# Lead Consolidation MVP – Slide Deck

This repository holds the code that builds `Lead_Consolidation_MVP.pptx` (24 slides).
All slide text, colours, and layout are in one file: `build.js`.

Do not edit the .pptx in PowerPoint. Every build replaces the .pptx with a new file from the code, so manual PowerPoint edits are lost.

## One-Time Setup

1. Install Node.js (LTS version) from https://nodejs.org
2. Install Git from https://git-scm.com
3. Clone the repository and install the packages:

```
git clone https://github.com/loosijie20-png/DXP-lead-consolidation.git
cd DXP-lead-consolidation
npm install
```

## Build the Deck

```
npm run build
```

The script creates `Lead_Consolidation_MVP.pptx` in this folder. The .pptx is not stored in Git (see `.gitignore`). Each person builds their own copy.

Close the .pptx in PowerPoint before you build. An open file cannot be replaced.

## Edit a Slide

Each slide in `build.js` starts with a comment. Search for the slide number:

| Slide | Comment in build.js |
|---|---|
| 1 | `// 1. Title` |
| 2 | `// 2. Problem & Pain Points` |
| 3 | `// 3. Proposed Solution Overview` |
| 4 | `// 4. Scope of the MVP` |
| 5 | `// 5. Three Approved Intake Methods` |
| 6 | `// 6. The Central Lead Register` |
| 7 | `// 7. Data Standardisation` |
| 8 | `// 8. Validation` |
| 9 | `// 9. Duplicate Detection` |
| 10 | `// 10. Lead Assignment` |
| 11 | `// 11. SLA Calculation` |
| 12 | `// 12. Notifications` |
| 13 | `// 13. Exception Management` |
| 14 | `// 14. Management Dashboard` |
| 15 | `// 15. Audit History` |
| 16 | `// 16. Technology Stack` |
| 17 | `// 17. Success Metrics` |
| 18 | `// 18. Delivery Approach` |
| 19 | `// 19. Testing & Quality Assurance` |
| 20 | `// 20. Documentation & Handover` |
| 21 | `// 21. Team Roles` |
| 22 | `// 22. Roles by Phase` |
| 23 | `// 23. Key Assumptions` |
| 24 | `// 24. Risks & Mitigations` |

- To change text: edit the text between the quotes. Keep the quotes and the commas.
- To change a colour for the whole deck: edit the `C` list near the top of the file.
- Slide order is the order of the blocks in the file. To move a slide, move its whole block.
- Footer page numbers update automatically.

## Team Workflow

Always get the latest version before you start:

```
git pull
```

After you edit and check the build:

```
git add build.js
git commit -m "Slide 11: change default SLA to 24 hours"
git push
```

Rules:

- Pull before you edit. Push soon after you finish.
- One commit per change. Name the slide number in the message.
- Tell the team in chat which slides you are editing. Two people editing the same slide at the same time causes a merge conflict.
- If `git push` is rejected, run `git pull`, then `git push` again.
- If `git pull` reports a conflict, open `build.js`, find the lines between `<<<<<<<` and `>>>>>>>`, keep the correct text, delete the markers, then commit and push.

## Files

| File | Purpose |
|---|---|
| `build.js` | All slide content and layout |
| `package.json` | Package list and the `build` command |
| `package-lock.json` | Exact package versions, so every build is the same |
| `.gitignore` | Keeps `node_modules` and the built .pptx out of Git |

## Tools Used

- [PptxGenJS](https://gitbrent.github.io/PptxGenJS/): writes the .pptx file
- [react-icons](https://react-icons.github.io/react-icons/) (Font Awesome 6 set): slide icons
- [sharp](https://sharp.pixelplumbing.com/): converts icons to images
