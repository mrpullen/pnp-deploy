import * as core from '@actions/core'

import { SPDefault } from '@pnp/nodejs'
import { SPFI, spfi } from '@pnp/sp'

import '@pnp/sp/webs/index'
import '@pnp/sp/appcatalog'
import { readFileSync } from 'fs'

export interface IPnPDeployOptions {
  siteUrl: string
  scopes: Array<string>
  tenantId: string
  clientId: string
  thumbprint: string
  packagePath: string
}

function getInputOptions(): IPnPDeployOptions {
  const requiredOption: core.InputOptions = {
    required: true,
    trimWhitespace: true
  }

  const options: IPnPDeployOptions = {
    siteUrl: core.getInput('siteUrl', requiredOption),
    scopes: core.getInput('scopes', requiredOption).split(','),
    tenantId: core.getInput('tenantId', requiredOption),
    clientId: core.getInput('clientId', requiredOption),
    thumbprint: core.getInput('thumbprint', requiredOption),
    packagePath: core.getInput('packagePath', requiredOption)
  }

  return options
}

function getCertBase64Encoded(): string {
  const requiredOption: core.InputOptions = {
    required: true,
    trimWhitespace: true
  };

  const certBase64Encoded = core.getInput('certBase64Encoded', requiredOption);
  core.setSecret(certBase64Encoded);

  return certBase64Encoded;
}

function initSPFI(options: IPnPDeployOptions, certBase64Encoded: string): SPFI {
  const sp = spfi().using(
    SPDefault({
      baseUrl: `${options.siteUrl}`,
      msal: {
        config: {
          auth: {
            authority: `https://login.microsoftonline.com/${options.tenantId}/`,
            clientId: options.clientId,
            clientCertificate: {
              thumbprint: options.thumbprint,
              privateKey: certBase64Encoded
            }
          }
        },
        scopes: options.scopes
      }
    })
  )
  return sp
}

function readPackageFile(options: IPnPDeployOptions): Buffer {
  try {
  const fileBuffer = readFileSync(options.packagePath)
  return fileBuffer
  }
  catch (error) {
    throw new Error(`Error reading file: (${options.packagePath}) - ${JSON.stringify(error, null, 4)}`)
  }
}

function getFileName(path: string): string {
  const parts = path.split('/')
  return parts[parts.length - 1]
}

/**
 * The main function for the action.
 *
 * @returns Resolves when the action is complete.
 */
export async function run(): Promise<void> {
  try {
    core.debug('Getting Input Options')
    const options = getInputOptions()

    core.debug(JSON.stringify(options, null, 4))
    const certBase64Encoded = getCertBase64Encoded();

    core.debug("CertBase64 - should be protected" + certBase64Encoded);
    core.debug('Configuring PnP SPFI');
    const sp = initSPFI(options, certBase64Encoded);

    const fileBuffer = readPackageFile(options)
    const fileName = getFileName(options.packagePath)
    const result = await sp.web.appcatalog.add(fileName, fileBuffer, true)

    core.setOutput('result', JSON.stringify(result, null, 4))
  } catch (error) {
    // Fail the workflow run if an error occurs
    if (error instanceof Error) core.setFailed(error.message)
  }
}
