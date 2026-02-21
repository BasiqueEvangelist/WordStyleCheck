- look into the relevant GOSTs
- add more lint types:
  * lints for bold/italic (e.g. in headers)
  * issues with lists
  * issues with tables
  * issues with images (e.g. captions)
  * issues with ToC
  * page numbers
  * page size and other section data
  * formulae????
- make sure existing lints don't fire on well formatted documents
  * restrict some lints to body text, others just to headers
  * don't fire lints on ToC (unless they are related to ToC stuff)
- figure out how to write out contexts for diagnostics better
- write GUI application (Avalonia or WinForms) for easier usage and step-by-step application of diagnostic autofixes