package com.zgh.java.tts;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class SpeechUtil {

	public void speakMessage(String message, int volume, int rate) {

		ActiveXComponent sap = new ActiveXComponent("Sapi.SpVoice");
		Dispatch sapo = sap.getObject();
		try {
			sap.setProperty("Volume", new Variant(volume));
			sap.setProperty("Rate", new Variant(rate));
			// sap.setProperty("AudioOutputStream", new Variant(rate));
			Dispatch.call(sapo, "Speak", new Variant(message));

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			sapo.safeRelease();
			sap.safeRelease();
		}
	}

	public static void main(String[] args) {
		SpeechUtil speechutil = new SpeechUtil();
		String message = "我们的使命与愿景成为世界级最专业、最具创新性的销售管理服务商!";
		speechutil.speakMessage(message, 100, 1);
	}
}