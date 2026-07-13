import type { OfficeHostIdentity } from './office-host-protocol';

export function officeHostIdentitiesEqual(left: OfficeHostIdentity, right: OfficeHostIdentity): boolean {
  return left.packageVersion === right.packageVersion
    && left.hostBuildId === right.hostBuildId
    && left.assetManifestDigest === right.assetManifestDigest;
}
