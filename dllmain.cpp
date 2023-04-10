#ifndef CEXPLORERCOMMAND_H
#define CEXPLORERCOMMAND_H

#include <Shobjidl.h> // Для IExplorerCommand и т.д.
#include <Shlwapi.h>
#pragma comment(lib, "shlwapi")
#include <wrl/module.h>
#include <wil/common.h>
#include <QString>

using namespace Microsoft::WRL;

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                       )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

class CExplorerCommand : public RuntimeClass<RuntimeClassFlags<ClassicCom>, IExplorerCommand>
{
public:
    CExplorerCommand() { };
    CExplorerCommand(QString appDir) : m_appDir() { };

    virtual const wchar_t* Title() = 0;
    virtual const EXPCMDFLAGS Flags() { return ECF_DEFAULT; }
    virtual const EXPCMDSTATE State(_In_opt_ IShellItemArray* selection) { return ECS_ENABLED; }

    // Inherited via IExplorerCommand
    HRESULT __stdcall GetTitle(_In_opt_ IShellItemArray* psiItemArray, LPWSTR* ppszName) override
    {
        SHStrDupW(L"WinPlugins", ppszName);
        return S_OK;
    }
    HRESULT __stdcall GetIcon(IShellItemArray* psiItemArray, LPWSTR* ppszIcon) override
    {
        QString iconPath = m_appDir + "WinPlugins.exe, -";
        SHStrDupW(iconPath.toStdWString().c_str(), ppszIcon);
        return S_OK;
    }
    HRESULT __stdcall GetToolTip(IShellItemArray* psiItemArray, LPWSTR* ppszInfotip) override
    {
        return E_NOTIMPL;
    }
    HRESULT __stdcall GetCanonicalName(GUID* pguidCommandName) override
    {
        *pguidCommandName = CLSID_NULL;
        return S_OK;
    }
    HRESULT __stdcall GetState(IShellItemArray* psiItemArray, BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState) override
    {
        *pCmdState = State(psiItemArray);
        return S_OK;
    }
    HRESULT __stdcall Invoke(IShellItemArray* psiItemArray, IBindCtx* pbc) override
    {
        return E_NOTIMPL;
    }
    HRESULT __stdcall GetFlags(EXPCMDFLAGS* pFlags) override
    {
        return E_NOTIMPL;
    }
    HRESULT __stdcall EnumSubCommands(IEnumExplorerCommand** ppEnum) override
    {
        return E_NOTIMPL;
    }

protected:
    QString m_appDir = "";
};

class __declspec(uuid("3282E233-C5D3-4533-9B25-44B8AAAFACFA")) TestExplorerCommandHandler final : public CExplorerCommand
{
public:
    const wchar_t* Title() override { return L"PhotoStore Command1"; }
    const EXPCMDSTATE State(_In_opt_ IShellItemArray* selection) override { return ECS_DISABLED; }
};

class __declspec(uuid("817CF159-A4B5-41C8-8E8D-0E23A6605395")) TestExplorerCommand2Handler final : public CExplorerCommand
{
public:
    const wchar_t* Title() override { return L"PhotoStore ExplorerCommand2"; }
};

class SubExplorerCommandHandler final : public CExplorerCommand
{
public:
    const wchar_t* Title() override { return L"SubCommand"; }
};

class CheckedSubExplorerCommandHandler final : public CExplorerCommand
{
public:
    const wchar_t* Title() override { return L"CheckedSubCommand"; }
    const EXPCMDSTATE State(_In_opt_ IShellItemArray* selection) override { return ECS_CHECKBOX | ECS_CHECKED; }
};

class RadioCheckedSubExplorerCommandHandler final : public CExplorerCommand
{
public:
    const wchar_t* Title() override { return L"RadioCheckedSubCommand"; }
    const EXPCMDSTATE State(_In_opt_ IShellItemArray* selection) override { return ECS_CHECKBOX | ECS_RADIOCHECK; }
};

class HiddenSubExplorerCommandHandler final : public CExplorerCommand
{
public:
    const wchar_t* Title() override { return L"HiddenSubCommand"; }
    const EXPCMDSTATE State(_In_opt_ IShellItemArray* selection) override { return ECS_HIDDEN; }
};

class EnumCommands : public RuntimeClass<RuntimeClassFlags<ClassicCom>, IEnumExplorerCommand>
{
public:
    EnumCommands()
    {
        m_commands.push_back(Make<SubExplorerCommandHandler>());
        m_commands.push_back(Make<CheckedSubExplorerCommandHandler>());
        m_commands.push_back(Make<RadioCheckedSubExplorerCommandHandler>());
        m_commands.push_back(Make<HiddenSubExplorerCommandHandler>());
        m_current = m_commands.cbegin();
    }

    // IEnumExplorerCommand
    IFACEMETHODIMP Next(ULONG celt, __out_ecount_part(celt, *pceltFetched) IExplorerCommand** apUICommand, __out_opt ULONG* pceltFetched)
    {
        ULONG fetched{ 0 };
        wil::assign_to_opt_param(pceltFetched, 0ul);

        for (ULONG i = 0; (i < celt) && (m_current != m_commands.cend()); i++)
        {
            m_current->CopyTo(&apUICommand[0]);
            m_current++;
            fetched++;
        }

        wil::assign_to_opt_param(pceltFetched, fetched);
        return (fetched == celt) ? S_OK : S_FALSE;
    }

    IFACEMETHODIMP Skip(ULONG /*celt*/) { return E_NOTIMPL; }
    IFACEMETHODIMP Reset()
    {
        m_current = m_commands.cbegin();
        return S_OK;
    }
    IFACEMETHODIMP Clone(__deref_out IEnumExplorerCommand** ppenum) { *ppenum = nullptr; return E_NOTIMPL; }

private:
    std::vector<ComPtr<IExplorerCommand>> m_commands;
    std::vector<ComPtr<IExplorerCommand>>::const_iterator m_current;
};

class __declspec(uuid("1476525B-BBC2-4D04-B175-7E7D72F3DFF8")) TestExplorerCommand3Handler final : public CExplorerCommand
{
public:
    const wchar_t* Title() override { return L"PhotoStore CommandWithSubCommands"; }
    const EXPCMDFLAGS Flags() override { return ECF_HASSUBCOMMANDS; }

    IFACEMETHODIMP EnumSubCommands(_COM_Outptr_ IEnumExplorerCommand** enumCommands) override
    {
        *enumCommands = nullptr;
        auto e = Make<EnumCommands>();
        return e->QueryInterface(IID_PPV_ARGS(enumCommands));
    }
};

class __declspec(uuid("30DEEDF6-63EA-4042-A7D8-0A9E1B17BB99")) TestExplorerCommand4Handler final : public CExplorerCommand
{
public:
    const wchar_t* Title() override { return L"PhotoStore Command4"; }
};

class __declspec(uuid("50419A05-F966-47BA-B22B-299A95492348")) TestExplorerHiddenCommandHandler final : public CExplorerCommand
{
public:
    const wchar_t* Title() override { return L"PhotoStore HiddenCommand"; }
    const EXPCMDSTATE State(_In_opt_ IShellItemArray* selection) override { return ECS_HIDDEN; }
};

/*CoCreatableClass(TestExplorerCommandHandler)
CoCreatableClass(TestExplorerCommand2Handler)
CoCreatableClass(TestExplorerCommand3Handler)
CoCreatableClass(TestExplorerCommand4Handler)
CoCreatableClass(TestExplorerHiddenCommandHandler)

CoCreatableClassWrlCreatorMapInclude(TestExplorerCommandHandler)
CoCreatableClassWrlCreatorMapInclude(TestExplorerCommand2Handler)
CoCreatableClassWrlCreatorMapInclude(TestExplorerCommand3Handler)
CoCreatableClassWrlCreatorMapInclude(TestExplorerCommand4Handler)
CoCreatableClassWrlCreatorMapInclude(TestExplorerHiddenCommandHandler)

STDAPI DllGetActivationFactory(_In_ HSTRING activatableClassId, _COM_Outptr_ IActivationFactory** factory)
{
    return Module<ModuleType::InProc>::GetModule().GetActivationFactory(activatableClassId, factory);
}

STDAPI DllCanUnloadNow()
{
    return Module<InProc>::GetModule().GetObjectCount() == 0 ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(_In_ REFCLSID rclsid, _In_ REFIID riid, _COM_Outptr_ void** instance)
{
    return Module<InProc>::GetModule().GetClassObject(rclsid, riid, instance);
}*/

#endif // CEXPLORERCOMMAND_H
