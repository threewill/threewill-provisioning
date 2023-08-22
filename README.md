# About this Repository

This repository is the home of ThreeWill's provisioning assets.  This includes the provisioning engine, which provides a customizable base for creating and provisioning a batch of sites including creating site collections, installing custom webparts, and invoking PnP Provisioning Templates.

For more information about using the provisioning engine, please refer to the doucmentation  
[Provisioning Engine Documentation](docs/index.md)

## Releasing a new version

Once all updates are complete including an approved Pull Request, you will need to tag the codebase with a new version. Ensure that this matches the version in `package-solution.json`. You may also find that you still need to increment the version in `package-solution.json` and, if that is the case, do that first then move onto this step.

```bash
git tag -a v1.x.x.x -m "Include a note about this release"
git push origin v1.x.x.x
```
