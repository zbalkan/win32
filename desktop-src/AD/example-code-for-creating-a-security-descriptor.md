---
title: Example Code for Creating a Security Descriptor
description: This topic includes PowerShell and C++ code examples that show how to create a security descriptor for an Active Directory object using ADSI.
ms.assetid: 7c6dcdaf-0bef-4f72-bd9d-dc3ab4295008
ms.tgt_platform: multiple
keywords:
- Example Code for Creating a Security Descriptor
ms.topic: reference
ms.date: 09/15/2026
topic_type: 
- kbArticle
api_name: 
api_type: 
api_location: 
---

# Example Code for Creating a Security Descriptor

The following examples show how to use Active Directory Service Interfaces (ADSI) to create a security descriptor for a new Active Directory object.

The examples create a new organizational unit, build a security descriptor with a discretionary access-control list (DACL), add an access-control entry (ACE), and assign the security descriptor before the new object is committed to Active Directory.

The caller must have permission to create the object in the parent container.

## PowerShell

> [!NOTE]
> For most Active Directory administration tasks in PowerShell, use the **ActiveDirectory PowerShell module**. It provides cmdlets for managing Active Directory objects, accounts, domains, forests, and related configuration. See [Active Directory module documentation](/powershell/module/activedirectory/about/about_activedirectory).
>
> ADSI provides lower-level access to Active Directory and is useful when the required operation is not directly exposed by the ActiveDirectory module or when direct access to ADSI interfaces is required.

```powershell
# ADSI constants.
$ADS_RIGHT_DS_READ_PROP             = 0x10
$ADS_ACETYPE_ACCESS_ALLOWED         = 0x00
$ADS_ACEFLAG_NONE                   = 0x00
$ADS_SD_CONTROL_SE_DACL_PRESENT     = 0x04
$ACL_REVISION_DS                    = 0x04

# Parent container and new object.
$parentPath   = 'LDAP://DC=Fabrikam,DC=com'
$relativeName = 'OU=Sales'

# Owner of the new object and trustee that receives the ACE.
$owner   = 'FABRIKAM\Administrator'
$trustee = 'FABRIKAM\Security Readers'

$container          = $null
$newObject          = $null
$securityDescriptor = $null
$dacl               = $null
$ace                = $null

try {
    # Bind to the parent container.
    $container =
        [System.Runtime.InteropServices.Marshal]::BindToMoniker($parentPath)

    # Create the directory object in the ADSI property cache.
    # The object is not committed until SetInfo is called.
    $newObject =
        $container.Create('organizationalUnit', $relativeName)

    # Create a new ADSI security descriptor.
    $securityDescriptor =
        New-Object -ComObject 'SecurityDescriptor'

    # Create a new DACL.
    $dacl =
        New-Object -ComObject 'AccessControlList'

    $dacl.AclRevision = $ACL_REVISION_DS

    # Create an ACE for the DACL.
    $ace =
        New-Object -ComObject 'AccessControlEntry'

    # Grant Read Property permission for all properties on this object.
    $ace.Trustee    = $trustee
    $ace.AccessMask = $ADS_RIGHT_DS_READ_PROP
    $ace.AceType    = $ADS_ACETYPE_ACCESS_ALLOWED
    $ace.AceFlags   = $ADS_ACEFLAG_NONE

    $dacl.AddAce($ace)

    # Configure the new security descriptor.
    $securityDescriptor.Revision = 1
    $securityDescriptor.Owner = $owner
    $securityDescriptor.Control =
        $ADS_SD_CONTROL_SE_DACL_PRESENT
    $securityDescriptor.DiscretionaryAcl = $dacl

    # Attach the security descriptor to the new object.
    $newObject.Put(
        'nTSecurityDescriptor',
        $securityDescriptor
    )

    # Create the object and commit the security descriptor.
    $newObject.SetInfo()
}
catch {
    throw "Unable to create '$relativeName': $($_.Exception.Message)"
}
finally {
    foreach ($comObject in @(
        $ace,
        $dacl,
        $securityDescriptor,
        $newObject,
        $container
    )) {
        if (
            $null -ne $comObject -and
            [System.Runtime.InteropServices.Marshal]::IsComObject($comObject)
        ) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $comObject
            )
        }
    }
}
```

## C++

The same ADSI interfaces can be accessed directly from C++. The following example performs the same operation as the PowerShell example.

```cpp
#include <windows.h>

#include <activeds.h>
#include <adshlp.h>
#include <atlbase.h>
#include <atlcomcli.h>
#include <stdio.h>

#pragma comment(lib, "Activeds.lib")
#pragma comment(lib, "Adsiid.lib")

HRESULT CreateOuWithSecurityDescriptor(
    const wchar_t* parent_path,
    const wchar_t* relative_name,
    const wchar_t* owner,
    const wchar_t* trustee) {
  if (parent_path == nullptr ||
      relative_name == nullptr ||
      owner == nullptr ||
      trustee == nullptr) {
    return E_INVALIDARG;
  }

  // Bind to the parent container using the caller's security context.
  CComPtr<IADsContainer> container;

  HRESULT hr = ADsOpenObject(
      parent_path,
      nullptr,
      nullptr,
      ADS_SECURE_AUTHENTICATION,
      IID_IADsContainer,
      reinterpret_cast<void**>(&container));

  if (FAILED(hr)) {
    return hr;
  }

  // Create the directory object in the ADSI property cache.
  CComPtr<IDispatch> object_dispatch;

  hr = container->Create(
      CComBSTR(L"organizationalUnit"),
      CComBSTR(relative_name),
      &object_dispatch);

  if (FAILED(hr)) {
    return hr;
  }

  CComQIPtr<IADs> new_object(object_dispatch);

  if (!new_object) {
    return E_NOINTERFACE;
  }

  // Create a new ADSI security descriptor.
  CComPtr<IADsSecurityDescriptor> security_descriptor;

  hr = CoCreateInstance(
      CLSID_SecurityDescriptor,
      nullptr,
      CLSCTX_INPROC_SERVER,
      IID_IADsSecurityDescriptor,
      reinterpret_cast<void**>(&security_descriptor));

  if (FAILED(hr)) {
    return hr;
  }

  // Create a new DACL.
  CComPtr<IADsAccessControlList> dacl;

  hr = CoCreateInstance(
      CLSID_AccessControlList,
      nullptr,
      CLSCTX_INPROC_SERVER,
      IID_IADsAccessControlList,
      reinterpret_cast<void**>(&dacl));

  if (FAILED(hr)) {
    return hr;
  }

  hr = dacl->put_AclRevision(ACL_REVISION_DS);

  if (FAILED(hr)) {
    return hr;
  }

  // Create a new ACE.
  CComPtr<IADsAccessControlEntry> ace;

  hr = CoCreateInstance(
      CLSID_AccessControlEntry,
      nullptr,
      CLSCTX_INPROC_SERVER,
      IID_IADsAccessControlEntry,
      reinterpret_cast<void**>(&ace));

  if (FAILED(hr)) {
    return hr;
  }

  // Grant Read Property permission for all properties on this object.
  hr = ace->put_Trustee(CComBSTR(trustee));

  if (FAILED(hr)) {
    return hr;
  }

  hr = ace->put_AccessMask(ADS_RIGHT_DS_READ_PROP);

  if (FAILED(hr)) {
    return hr;
  }

  hr = ace->put_AceType(ADS_ACETYPE_ACCESS_ALLOWED);

  if (FAILED(hr)) {
    return hr;
  }

  // Apply the ACE only to the new object.
  hr = ace->put_AceFlags(0);

  if (FAILED(hr)) {
    return hr;
  }

  CComQIPtr<IDispatch> ace_dispatch(ace);

  if (!ace_dispatch) {
    return E_NOINTERFACE;
  }

  hr = dacl->AddAce(ace_dispatch);

  if (FAILED(hr)) {
    return hr;
  }

  CComQIPtr<IDispatch> dacl_dispatch(dacl);

  if (!dacl_dispatch) {
    return E_NOINTERFACE;
  }

  // Configure the new security descriptor.
  hr = security_descriptor->put_Revision(1);

  if (FAILED(hr)) {
    return hr;
  }

  hr = security_descriptor->put_Owner(
      CComBSTR(owner));

  if (FAILED(hr)) {
    return hr;
  }

  hr = security_descriptor->put_Control(
      ADS_SD_CONTROL_SE_DACL_PRESENT);

  if (FAILED(hr)) {
    return hr;
  }

  hr = security_descriptor->put_DiscretionaryAcl(
      dacl_dispatch);

  if (FAILED(hr)) {
    return hr;
  }

  // Put requires the security descriptor as an IDispatch VARIANT.
  CComQIPtr<IDispatch> security_descriptor_dispatch(
      security_descriptor);

  if (!security_descriptor_dispatch) {
    return E_NOINTERFACE;
  }

  CComVariant security_descriptor_value(
      security_descriptor_dispatch);

  hr = new_object->Put(
      CComBSTR(L"nTSecurityDescriptor"),
      security_descriptor_value);

  if (FAILED(hr)) {
    return hr;
  }

  // Create the object and commit the security descriptor.
  return new_object->SetInfo();
}

int wmain() {
  HRESULT hr = CoInitializeEx(
      nullptr,
      COINIT_APARTMENTTHREADED);

  if (FAILED(hr)) {
    fwprintf(
        stderr,
        L"COM initialization failed: 0x%08lX\n",
        static_cast<unsigned long>(hr));

    return 1;
  }

  hr = CreateOuWithSecurityDescriptor(
      L"LDAP://DC=Fabrikam,DC=com",
      L"OU=Sales",
      L"FABRIKAM\\Administrator",
      L"FABRIKAM\\Security Readers");

  CoUninitialize();

  if (FAILED(hr)) {
    fwprintf(
        stderr,
        L"Unable to create the directory object: 0x%08lX\n",
        static_cast<unsigned long>(hr));

    return 1;
  }

  return 0;
}
```

The C++ example uses `IADsContainer::Create` to prepare the new directory object in the ADSI property cache. It then creates `CLSID_SecurityDescriptor`, `CLSID_AccessControlList`, and `CLSID_AccessControlEntry` COM objects and assigns the completed security descriptor before calling `IADs::SetInfo`.

For more information, see [Creating Security Descriptors for New Directory Objects](creating-a-security-descriptor-for-a-new-directory-object.md), [IADsContainer::Create](/windows/win32/api/iads/nf-iads-iadscontainer-create), [IADsSecurityDescriptor](/windows/win32/api/iads/nn-iads-iadssecuritydescriptor), and [IADsAccessControlEntry](/windows/win32/api/iads/nn-iads-iadsaccesscontrolentry).

## Security Descriptor Contents

The examples create a security descriptor with an explicit owner and DACL.

The DACL contains one explicit allow ACE that grants `ADS_RIGHT_DS_READ_PROP` to the trustee. Because no object type GUID is specified, the permission applies to all properties of the new object.

The examples do not create a SACL.

## Inheritance

The examples set `ADS_SD_CONTROL_SE_DACL_PRESENT` but do not set `ADS_SD_CONTROL_SE_DACL_PROTECTED`.

When a security descriptor is explicitly supplied during creation of an Active Directory object, Active Directory Domain Services merges inheritable ACEs from the parent into the new object's DACL unless the DACL is protected.

See [How Security Descriptors are Set on New Directory Objects](how-security-descriptors-are-set-on-new-directory-objects.md).

## ACE Ordering

ACE order affects access evaluation. Explicit deny ACEs normally precede explicit allow ACEs, and explicit ACEs precede inherited ACEs.

The examples create only one explicit ACE. Applications that construct DACLs containing multiple ACEs should add them in canonical order.

See [Order of ACEs in a DACL](../SecAuthZ/order-of-aces-in-a-dacl.md).

## Existing Objects

A newly created ADSI security descriptor can also be assigned to the `nTSecurityDescriptor` property of an existing object. Doing so replaces the existing security descriptor information being written.

For routine permission changes on existing objects, retrieve the current security descriptor and modify its DACL instead of constructing a replacement descriptor. This preserves unrelated access-control information.

See [Setting Access Rights on an Object](setting-access-rights-on-an-object.md).

## Empty and NULL DACLs

An empty DACL contains no ACEs and therefore grants no discretionary access.

A NULL DACL has different semantics and permits unrestricted access.

Code that creates a security descriptor must distinguish carefully between these two cases.
