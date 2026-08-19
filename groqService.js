// groqService.js
const GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions";

async function tanyaAI(question, contextData = {}) {
  const apiKey = process.env.GROQ_API_KEY;
  if (!apiKey) {
    throw new Error("GROQ_API_KEY belum dikonfigurasi di server");
  }

  const systemPrompt = `Kamu adalah asisten AI untuk aplikasi SIMAMORA BPS (Sistem Monitoring Kinerja).
Jawablah pertanyaan seputar data kinerja tim, realisasi kerja, kendala, solusi, dan RTL berdasarkan data berikut.
Jawab singkat, jelas, dalam Bahasa Indonesia.

Data konteks saat ini:
${JSON.stringify(contextData, null, 2)}`;

  const response = await fetch(GROQ_API_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: "openai/gpt-oss-120b",
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: question },
      ],
      temperature: 0.4,
      max_tokens: 1024,
    }),
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Groq API error (${response.status}): ${errText}`);
  }

  const data = await response.json();
  const answer = data?.choices?.[0]?.message?.content;
  if (!answer) throw new Error("Groq API tidak mengembalikan jawaban yang valid");
  return answer;
}

module.exports = { tanyaAI };