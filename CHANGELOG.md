# Changelog

All notable changes to this project will be documented in this file. See [standard-version](https://github.com/conventional-changelog/standard-version) for commit guidelines.

## [2.3.0](https://github.com/tks18/xl-pq-handler/compare/v2.2.0...v2.3.0) (2025-10-28)


### Bug Fixes 🛠

* **components/codeviewer:** remove line numbers as rendering in tkinter is gimmicky ([836d1c6](https://github.com/tks18/xl-pq-handler/commit/836d1c64b697a4fd5ca65ec0c48caf2184333cb0))


### Features 🔥

* **ui/views:** write a component to ui/library to refactor the treeview part ([1fe49fc](https://github.com/tks18/xl-pq-handler/commit/1fe49fc54542cf85d3f432deb8b4340580f08721))
* **ui/views:** write a new component for ui/extract to refactor file extraction part ([b34341e](https://github.com/tks18/xl-pq-handler/commit/b34341ec243159a7184d27e3b3b59a1021213334))
* **ui/views:** write a new component for ui/extract to refactor log viewer part ([c5050e7](https://github.com/tks18/xl-pq-handler/commit/c5050e7ff2ff398321dc8117d2e8ffa39c52b603))
* **ui/views:** write a new component for ui/extract to refactor workbook extraction part ([d0bd020](https://github.com/tks18/xl-pq-handler/commit/d0bd02002078b692b428f5dfb82f34047192cace))
* **ui/views:** write a new component for ui/library for refactoring top bar part ([eb9cce9](https://github.com/tks18/xl-pq-handler/commit/eb9cce93865b5e743559ad91309d247a7316f6b2))
* **ui/views:** write a new component to refactor bottom panel ([6a69c84](https://github.com/tks18/xl-pq-handler/commit/6a69c840b2643cb6102748e6c0efaff8e179c77d))
* **ui/views:** write the edit dialog as a seperate component to refactor ([956cf34](https://github.com/tks18/xl-pq-handler/commit/956cf34a69c396dd48d8ae8c70e51edc805dbf6c))


### Code Refactoring 🖌

* **ui/views:** refactor the create view to a particular folder for better management ([9f81259](https://github.com/tks18/xl-pq-handler/commit/9f81259fe5cd5aead1a0ac3137bc5ffffc6cfc36))
* **ui/views:** refactor the ui/extract completely to use the new components ([5e88106](https://github.com/tks18/xl-pq-handler/commit/5e8810663074999d6f557fbb66283bea1b72c984))
* **ui/views:** refactor the ui/library completely to use the new components ([e1b2e1d](https://github.com/tks18/xl-pq-handler/commit/e1b2e1d6f98b3e6cfed74cce7bb0fccf180df3db))
* **ui/views:** remove the extract.py file as it is refactored ([b1ad98f](https://github.com/tks18/xl-pq-handler/commit/b1ad98fcd2f07defad81e935b1a1b200096ccc52))


### Others 🔧

* update version ([3a1e5ff](https://github.com/tks18/xl-pq-handler/commit/3a1e5ff752221f3edf70612db15d2b80aeb4350f))

## [2.2.0](https://github.com/tks18/xl-pq-handler/compare/v2.1.1...v2.2.0) (2025-10-28)


### Features 🔥

* **api:** introducing a regex based m code parser ([9e0a1e4](https://github.com/tks18/xl-pq-handler/commit/9e0a1e4b5fd7fef59ed8b7f927d2779ffb434f01))
* **api:** update the manager to expose the functions to use the parser ([925bd36](https://github.com/tks18/xl-pq-handler/commit/925bd36e4999522de13733460d40d359ea725483))
* **ui:** update views - library, extract & create to use the new component for rendering preview ([19c2234](https://github.com/tks18/xl-pq-handler/commit/19c223462daea426c3501b75ee2ad8040486c3d8))
* **ui:** write a new component - codeview to enhance textbox widget ([962596a](https://github.com/tks18/xl-pq-handler/commit/962596a268e3a69abd8e73db99809c18ee662d9b))


### Others 🔧

* update pyproject.toml version ([c28ad60](https://github.com/tks18/xl-pq-handler/commit/c28ad60b2f72c36f34117381e1d703c1d4feaae2))

### [2.1.1](https://github.com/tks18/xl-pq-handler/compare/v2.1.0...v2.1.1) (2025-10-27)


### Bug Fixes 🛠

* **paths:** update the paths everywhere to save the scripts under functions folder only ([b0bfb35](https://github.com/tks18/xl-pq-handler/commit/b0bfb3559bda8279d801b4731598bd900a45f6a8))


### Others 🔧

* update pyproject.toml file version ([1150ecb](https://github.com/tks18/xl-pq-handler/commit/1150ecb800ceebe307a4891143709ce7a7ece3b4))

## [2.1.0](https://github.com/tks18/xl-pq-handler/compare/v2.0.3...v2.1.0) (2025-10-27)


### Features 🔥

* **ui/extract:** feature: allow extraction from multiple open workbooks ([84f2a7c](https://github.com/tks18/xl-pq-handler/commit/84f2a7c3eae8e885c9ee15ad4d80d5ba3e817cdf))
* **ui/library:** allow insertion into multiple wbs not only active, add context menu to edit ([908713c](https://github.com/tks18/xl-pq-handler/commit/908713c033b20b82edcf96cab2bf9d43df3a28a9))
* **ui:** update the entire ui & manager to support the new changes ([3109a5b](https://github.com/tks18/xl-pq-handler/commit/3109a5bd68be9fe46140601e9a0a538d69d1b90e))


### Docs 📃

* **readme:** update readme, bump to version 2.1.0 ([9079349](https://github.com/tks18/xl-pq-handler/commit/907934980db99cc6e8c6400eb639b129e7807819))

### [2.0.3](https://github.com/tks18/xl-pq-handler/compare/v2.0.2...v2.0.3) (2025-10-26)


### Others 🔧

* make the script runnable by renaming the file to __main__ ([81efb1d](https://github.com/tks18/xl-pq-handler/commit/81efb1dc2a75e2f9fef1c3ff526a7099e3ea3428))

### [2.0.2](https://github.com/tks18/xl-pq-handler/compare/v2.0.1...v2.0.2) (2025-10-26)


### Code Refactoring 🖌

* **app:** change the title properly ([917155e](https://github.com/tks18/xl-pq-handler/commit/917155edbf9c036069dd7c24b9a356a070116401))


### Others 🔧

* update version ([fc98186](https://github.com/tks18/xl-pq-handler/commit/fc981863e62e568e267ff0c5c1be2f2e44c04d6d))
* update version ([0b3400f](https://github.com/tks18/xl-pq-handler/commit/0b3400fa8d59974b9033ad19cc6ffb19a6fbbc1c))

### [2.0.1](https://github.com/tks18/xl-pq-handler/compare/v2.0.0...v2.0.1) (2025-10-26)


### Others 🔧

* update pyproject.toml version ([3e88e67](https://github.com/tks18/xl-pq-handler/commit/3e88e670b695ee48194b273f2cd7b46534fe665f))


### Bug Fixes 🛠

* **app:** remove unwanted import which caused problems ([6e24830](https://github.com/tks18/xl-pq-handler/commit/6e24830f9aea851aa383dcbc6c588c4053eee1d5))

## [2.0.0](https://github.com/tks18/xl-pq-handler/compare/v1.1.1...v2.0.0) (2025-10-26)


### CI 🛠

* **docs:** add commitlint, husky ci step to standardize commit messages & versioning ([fe05889](https://github.com/tks18/xl-pq-handler/commit/fe058891fd41a5a95579752f6b1cf77446a7ad96))


### Code Refactoring 🖌

* **class/handler:** remove the monolith handler file and refactor it completely (future) ([b1b9518](https://github.com/tks18/xl-pq-handler/commit/b1b951800fdff806fb13b4d792c0a1277daf0414))


### Features 🔥

* **app:** create a cli app that runs the UI from cmdline ([10388c4](https://github.com/tks18/xl-pq-handler/commit/10388c493fd3dccb601e73ab0b4f740d548af73a))
* **class/dependencies:** add a dependencyhandler class ([97c6859](https://github.com/tks18/xl-pq-handler/commit/97c685963e2f99c291f2e49056d4ada4e6a815c3))
* **class/excel_service:** write the excel_service class along with utils ([b3cd481](https://github.com/tks18/xl-pq-handler/commit/b3cd48198bf5725d301d4bd2555147971fe36502))
* **class/models:** write a new class to define the model of powerquery file ([d908ad7](https://github.com/tks18/xl-pq-handler/commit/d908ad760982bdb38b78d9c29787bf714c2d00c7))
* **class/pqmanager:** orchestrate the entire manager class ([0b6f468](https://github.com/tks18/xl-pq-handler/commit/0b6f468e816a04e936e4ee6002c99629009f5d03))
* **class/storage:** write a storage class to handle all file operations ([a5cd209](https://github.com/tks18/xl-pq-handler/commit/a5cd2090dc80ec5dbe938a7744fd43b96c4e9a75))
* **ui:** move the entire UI from another repo modularizing it ([0d2b4e2](https://github.com/tks18/xl-pq-handler/commit/0d2b4e23c1b213774839cf562194fdccf83faa27))
