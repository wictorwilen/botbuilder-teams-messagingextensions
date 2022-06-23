# Change Log

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/)
and this project adheres to [Semantic Versioning](http://semver.org/).



## [*1.8.1*] - <*2022-06-23*>

### Fixes

* Updated dependencies

## [*1.8.0*] - <*2021-10-22*>

### Added

* Added support for Adaptive Cards 1.4 universal actions `adaptiveCard/action` (#12)
* Added unit tests - 100% code coverage

### Changes

* Moved from Travis CI to Github Actions
* Migrated from TSLint to ESLint
* Migrated to botbuilder 4.14.1

### Fixes

* Fixed issues where submitAction could return an error if overrides was not defined
* Fixed an issue where selectItem returned a success even when an error occurred

## [*1.7.0*] - <*2020-10-27*>

### Changes

* Using bot framework `^4.9.0`

### Added

* Support for `send` and `edit` for `submitActions` (#10)

## [*1.6.0*] - <*2020-05-17*>

### Fixed

* Fixed an issue where `context.activity.value` is undefined (#6, #7)

### Changed

* Moved `botbuilder-core` to `devDependencies`

### Removed

* Removed `ms-rest-js` package (#1)

## [*1.5.0*] - <*2020-03-05*>

### Added

* Added logging (`msteams` namespace)

### Fixed

* Fixed the `type` of the response to `invokeResponse`

### Changed

* Migrated to `botbuilder-core@4.7.1`
* Breaking changes in the `IMessagingExtensionMiddlewareProcessor` where
types from `botbuilder-core` is used instead of custom defintions
* Updated Travis build settings

### Removed

* Removed `botbuilder-teams`
* Removed all custom interface declarations

## [*1.4.0*]- <*2019-06-02*>

### Changed

* `onQueryLink` is no longer filtering on `commandId` (as that is not sent in the `composeExtension/onQueryLink`)
* Changed signature for `onQueryLink` to use `IAppBasedLinkQuery` as value, to match official swagger
* Updated devDependencies

## [*1.3.0*] - <*2019-05-22*>

### Changed

* Changed signature for `onSubmitAction` and `onFetchTask` to use new `IMessagingExtensionActionRequest` interface
* Changed signature for `onFetchTask` to return `MessagingExtensionResult` (for `auth` and `config`) or `ITaskModuleResult` (when using `continue` or `message`)

## [*1.2.1*] - <*2019-05-07*>

### Changed
* Fixed versions for dependencies

## [*1.2.0*] - <*2019-05-06*>

### Added
* Added support for action command responses (`onFetchTask` - `composeExtension/fetchTask`)
* Added support for `Action.Submit` from adaptive cards (`onCardButtonClicked` - `composeExtension/onCardButtonClicked`)
* Added support for select item in Message Extensions (`onSelectItem` - `composeExtension/selectItem`)

## [*1.1.0*] - <*2019-04-29*>

### Added
* Added support for Link unfurling (`onQueryLink` - `composeExtension/queryLink`)
* Added support for Message Actions (`onSubmitAction` - `composeExtension/submitAction`)

### Changes
* Made all methods of `IMessagingExtensionMiddlewareProcessor` optional

## [*1.0.0*] - <*2019-03-29*>

### Added
* Initial release