import * as core from '@actions/core'
import { SPFI, spfi } from '@pnp/sp'
import '@pnp/sp/webs'
import '@pnp/sp/appcatalog'
import { SPDefault } from '@pnp/nodejs'
import { readFileSync } from 'fs'
export interface IPnPDeployOptions {
  siteUrl: string
  scopes: Array<string>
  tenantId: string
  clientId: string
  thumbprint: string
  packagePath: string
}

const getInputOptions = (): IPnPDeployOptions => {
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

const getCertBase64Encoded = (): string => {
  const requiredOption: core.InputOptions = {
    required: true,
    trimWhitespace: true
  }

  const certBase64Encoded = core.getInput('certBase64Encoded', requiredOption)
  core.setSecret(certBase64Encoded)

  return certBase64Encoded
}

const initSPFI = (
  options: IPnPDeployOptions,
  certBase64Encoded: string
): SPFI => {
  core.debug('Initializing SPFI')
  core.debug(`SiteUrl: ${options.siteUrl}`)
  core.debug(`Scopes: ${options.scopes}`)
  core.debug(`TenantId: ${options.tenantId}`)
  core.debug(`ClientId: ${options.clientId}`)
  core.debug(`Thumbprint: ${options.thumbprint}`)
  core.debug(
    `Authority: https://login.microsoftonline.com/${options.tenantId}/`
  )

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

const readPackageFile = (options: IPnPDeployOptions): Buffer => {
  try {
    const fileBuffer = readFileSync(options.packagePath)
    return fileBuffer
  } catch (error) {
    throw new Error(
      `Error reading file: (${options.packagePath}) - ${JSON.stringify(error, null, 4)}`
    )
  }
}

const getFileName = (path: string): string => {
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
    const certBase64Encoded = getCertBase64Encoded()

    core.debug('CertBase64 - should be protected' + certBase64Encoded)
    core.debug('Configuring PnP SPFI')
    const sp = initSPFI(options, certBase64Encoded)
    core.debug('Reading Package File')
    const fileBuffer = readPackageFile(options)
    core.debug('Retrieving File Name')
    const fileName = getFileName(options.packagePath)

    core.debug(`Adding File to App Catalog ${fileName}`)
    const result = await sp.web.appcatalog.add(fileName, fileBuffer, true)

    core.setOutput('result', JSON.stringify(result, null, 4))
  } catch (error) {
    // Fail the workflow run if an error occurs
    if (error instanceof Error) core.setFailed(error.message)
  }
}
