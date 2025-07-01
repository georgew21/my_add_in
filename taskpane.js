Office.initialize = () => {};

async function sendPrompt() {
  const my_prompt = document.getElementById("promptInput").value;
  const status = document.getElementById("status");

  Office.context.mailbox.item.body.getAsync("text", async (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      status.innerText = "❌ Αποτυχία ανάγνωσης email.";
      return;
    }

    const mail_to_respond = result.value;

    const payload = {
      my_prompt,
      mail_to_respond
    };

    try {
      const response = await fetch("https://hook.eu2.make.com/hqsxuqhay7lfvod4pctepuc55blwj1uv", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });

      if (response.ok) {
        status.innerText = "✅ Εστάλη επιτυχώς!";
      } else {
        status.innerText = `❌ Σφάλμα: ${response.status}`;
      }
    } catch (err) {
      status.innerText = "❌ Σφάλμα σύνδεσης στο webhook.";
    }
  });
}
