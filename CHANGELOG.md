# Change Log

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## [*Unreleased*]

### Changed

* `onQueryLink` is no longer filtering on `commandId` (as that is not sent in the `composeExtension/onQueryLink`)
* Changed signature for `onQueryLink` to use `IAppBasedLinkQuery` as value, to match official swagger 

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