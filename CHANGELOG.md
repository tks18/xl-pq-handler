# Changelog

All notable changes to this project will be documented in this file. See [standard-version](https://github.com/conventional-changelog/standard-version) for commit guidelines.

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
