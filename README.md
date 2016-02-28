# About the tool
More info here: http://mikisoft.me/programs/unibot

**Original author's note for the main repository: The project is abandoned so it's highly unlikely that I'll be doing further work on it. Also, mind that pull requests will be automatically rejected!**

# Differences between v1.4 and this version
- fixed **a ton** of bugs
- added **index manager**
- added in _Post field_ **DELETE** (-), **HEAD** (@) and **PUT** (\<file.ext>) support
- added **Minimize to tray** and **Always on top** options in _EXE bot_
- added new dependent commands (functions):
  - **dech()** - decode _HTML entities_
  - **num()** - convert and treat string as numeric value
  - **b64d()** - _Base64_ decode
- **Simple Captcha** plugin updated to _v1.2_ and **captcha9kw** plugin updated to _v1.1_, added in both:
  - _User-Agent_ and _Referer_ header
  - _(Static) GIF_ support
  - Settings dialog with more options
- **GSA Captcha Breaker** plugin updated to _v1.1_, added:
  - _User-Agent_ and _Referer_ header
  - _GIF_ support
- **2captcha** plugin included and updated to _v1.1_, added:
  - Ability to report bad captcha text
- **UniBot plugin utility** also included in the package
- plugin interface has changed (along with other changes, _BuildSettings_ procedure is added)
- problem with _DEP_ solved

# ToDo
- fix the existing bug with public strings, arrays (race condition) and _TrimO_ function, and then update the _EXE bot_ code according to the main project
- finish wizard (with _X-CSRF-Token_ detection and _basic HTTP auth_)
- make status bar more active in the process
- update the walkthrough

# Feature suggestions
- **Don't rotate proxy** option (in _Proxy and thread settings_ dialog)
- independent commands:
  - **[dinp]** - delayed input, which unlike _[inp]_ will show once it's reached to it in the process
  - **[(e)tstamp]** - timestamp of current time; e - EPOCH; with the feature of adding time to it, e.g: **[tstamp+10m]** (which would return timestamp of 10 minutes in the future)
- Regex parameter for concatenating/addition of all occurences
- **if()** command and grouped clauses

# Donation
Go here: http://mikisoft.me/uncategorized/bitcoin-donation
