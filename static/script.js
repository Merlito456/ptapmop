async function pasteImage(target) {
  try {
    const items = await navigator.clipboard.read();

    for (const item of items) {
      for (const type of item.types) {
        if (type.startsWith("image/")) {
          const blob = await item.getType(type);
          const reader = new FileReader();

          reader.onload = function(event) {
            const dataUrl = event.target.result;
            document.getElementById(`${target}_image_paste`).value = dataUrl;
            document.getElementById(`${target}_preview`).src = dataUrl;
          };

          reader.readAsDataURL(blob);
          return;
        }
      }
    }

    alert("No image found in clipboard.");
  } catch (err) {
    alert("Clipboard paste failed. Your browser may block clipboard access.");
    console.error(err);
  }
}