using System.Speech.Synthesis;

namespace AgoraIn.ClassIslandPlugin.Services;

/// <summary>
/// 呼叫语音朗读（Windows 中文 TTS，System.Speech）
/// 独立于提醒 UI：同一文本不重复朗读；无 TTS 环境时静默降级
/// </summary>
public static class TtsSpeaker
{
    /// <summary>最近一次朗读的话语（避免重复轮询/重复呼叫时重读）</summary>
    private static string? _lastSpoken;

    public static void Speak(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return;
        if (text == _lastSpoken) return;
        _lastSpoken = text;

        Task.Run(() =>
        {
            try
            {
                using var synth = new SpeechSynthesizer();
                var zh = synth.GetInstalledVoices()
                    .FirstOrDefault(v => v.VoiceInfo.Culture.Name.StartsWith("zh", StringComparison.OrdinalIgnoreCase))
                    ?.VoiceInfo.Name;
                if (!string.IsNullOrEmpty(zh)) synth.SelectVoice(zh);
                synth.Rate = 1;
                synth.Speak(text);
            }
            catch
            {
                // 无 TTS 环境时静默降级：仅显示提醒
            }
        });
    }
}
