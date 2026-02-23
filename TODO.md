- add more lint types:
  * issues with lists
    + reconstruct handmade lists into actual proper lists
    + adjust lists to have proper properties
    + !!! do not conflict with the needless paragraph break lint
  * issues with tables
  * issues with images (e.g. captions)
  * issues with ToC
  * page numbers
  * page size and other section data
  * formulae????
- make lint autofixers able to take in whether to generate revisions or not
- write autoclassifier for captions that don't have a style named "Caption" in the inheritance tree
- write GUI application (Avalonia or WinForms) for easier usage and step-by-step application of diagnostic autofixes