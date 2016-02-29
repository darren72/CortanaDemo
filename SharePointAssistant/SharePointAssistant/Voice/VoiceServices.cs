using System;
using Windows.ApplicationModel.VoiceCommands;
using Windows.Storage;

namespace SharePointAssistant.Voice
{
    /// <summary>
    /// Voice processing services.
    /// </summary>
    public static class VoiceServices 
    {
        /// <summary>
        /// Registers the available voice commands with the application.
        /// </summary>
        public static async void RegisterVoiceCommands()
        {
            var storageFile = await StorageFile.GetFileFromApplicationUriAsync(new Uri("ms-appx:///VoiceCommands.xml"));
            await VoiceCommandDefinitionManager.InstallCommandDefinitionsFromStorageFileAsync(storageFile);
        }
    }
}
