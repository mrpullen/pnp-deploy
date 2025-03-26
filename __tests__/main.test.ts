/* eslint-disable @typescript-eslint/no-explicit-any */
import { run } from '../src/main'
import * as core from '@actions/core'
import * as fs from 'fs'
import { SPFI } from '@pnp/sp'

jest.mock('@actions/core')
jest.mock('fs')
jest.mock('@pnp/sp')

describe('run', () => {
  beforeEach(() => {
    jest.clearAllMocks()
  })

  it('should successfully execute the action', async () => {
    const mockGetInput = core.getInput as jest.Mock
    const mockSetOutput = core.setOutput as jest.Mock
    const mockDebug = core.debug as jest.Mock
    const mockSetFailed = core.setFailed as jest.Mock
    const mockReadFileSync = fs.readFileSync as jest.Mock
    const mockAdd = jest.fn()

    mockGetInput.mockImplementation((name: string) => {
      const inputs: Record<string, string> = {
        siteUrl: 'https://example.sharepoint.com',
        scopes: 'scope1,scope2',
        tenantId: 'tenant-id',
        clientId: 'client-id',
        thumbprint: 'thumbprint',
        packagePath: '/path/to/package.sppkg',
        certBase64Encoded: 'mock-cert-base64'
      }
      return inputs[name]
    })

    mockReadFileSync.mockReturnValue(Buffer.from('mock file content'))
    ;(SPFI.prototype.web as any) = {
      appcatalog: {
        add: mockAdd.mockResolvedValue({ success: true })
      }
    }

    await run()

    expect(mockDebug).toHaveBeenCalledWith('Getting Input Options')
    expect(mockDebug).toHaveBeenCalledWith(
      JSON.stringify(
        {
          siteUrl: 'https://example.sharepoint.com',
          scopes: ['scope1', 'scope2'],
          tenantId: 'tenant-id',
          clientId: 'client-id',
          thumbprint: 'thumbprint',
          packagePath: '/path/to/package.sppkg'
        },
        null,
        4
      )
    )
    expect(mockDebug).toHaveBeenCalledWith(
      'CertBase64 - should be protectedmock-cert-base64'
    )
    expect(mockDebug).toHaveBeenCalledWith('Configuring PnP SPFI')
    expect(mockAdd).toHaveBeenCalledWith(
      'package.sppkg',
      Buffer.from('mock file content'),
      true
    )
    expect(mockSetOutput).toHaveBeenCalledWith(
      'result',
      JSON.stringify({ success: true }, null, 4)
    )
    expect(mockSetFailed).not.toHaveBeenCalled()
  })

  it('should fail the action if an error occurs', async () => {
    const mockSetFailed = core.setFailed as jest.Mock
    const mockGetInput = core.getInput as jest.Mock

    mockGetInput.mockImplementation(() => {
      throw new Error('Mock error')
    })

    await run()

    expect(mockSetFailed).toHaveBeenCalledWith('Mock error')
  })
})
