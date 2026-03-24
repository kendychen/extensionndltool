// Inject trước khi trang load — vô hiệu hóa passkey/fido
// để Microsoft hiện form nhập password thay vì passkey prompt
delete navigator.credentials;
window.PublicKeyCredential = undefined;
