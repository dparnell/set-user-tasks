use windows::{
    core::{PCWSTR, Interface},
    Win32::{
        System::{
            Com::{CoCreateInstance, CLSCTX_INPROC_SERVER}
        },
        UI::{
            Shell::{ICustomDestinationList, IShellLinkW, Common::{IObjectCollection, IObjectArray}, DestinationList, EnumerableObjectCollection, ShellLink, PropertiesSystem::{IPropertyStore, InitPropVariantFromStringVector}}
        },
    }
};
use windows::Win32::Storage::EnhancedStorage::PKEY_Title;


/// Construct a `windows-rs`'s [`PWSTR`] from a [`&str`].
///
/// See: https://github.com/microsoft/windows-rs/issues/973#issue-942298423
#[inline]
fn string_to_pwstr(str: &str) -> PCWSTR {
    let mut encoded = str.encode_utf16().chain([0u16]).collect::<Vec<u16>>();
    PCWSTR::from_raw(encoded.as_mut_ptr())
}

pub struct UserTask<'a> {
    pub title: &'a str,
    pub description: &'a str,
    pub arguments: &'a str,
    pub icon_path: &'a str,
    pub icon_index: i32,
    pub program: &'a str,
}

pub fn set_tasks(tasks: Vec<UserTask>) -> Result<(), windows::core::Error> {
    let cdl: ICustomDestinationList = unsafe {
        CoCreateInstance(&DestinationList, None, CLSCTX_INPROC_SERVER)?
    };

    let collection: IObjectCollection = unsafe {
        CoCreateInstance(&EnumerableObjectCollection, None, CLSCTX_INPROC_SERVER)?
    };

    unsafe {
        for task in tasks {
            let shell_link: IShellLinkW = CoCreateInstance(&ShellLink, None, CLSCTX_INPROC_SERVER)?;

            shell_link.SetDescription(string_to_pwstr(task.description))?;
            shell_link.SetArguments(string_to_pwstr(task.arguments))?;
            shell_link.SetIconLocation(string_to_pwstr(task.icon_path), task.icon_index)?;
            shell_link.SetPath(string_to_pwstr(task.program)  )?;

            let prop_store: IPropertyStore = shell_link.cast()?;
            let pkey = PKEY_Title;

            let mut title: Vec<u16> = task.title.encode_utf16().collect();
            title.push(0x00);
            let pv = InitPropVariantFromStringVector(&[windows::core::PWSTR(title.as_mut_ptr())])?;

            prop_store.SetValue(&pkey, &pv)?;
            prop_store.Commit()?;

            collection.AddObject(&shell_link)?;
        }

        let mut slots_visible: u32 = 0;
        let _removed: IObjectArray = cdl.BeginList(&mut slots_visible as *mut u32)?;

        cdl.AddUserTasks(&collection)?;
        cdl.CommitList()?;
    }

    Ok(())
}