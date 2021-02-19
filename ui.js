function displayProfile(user) {    
    if (!user) {
        return;
    }

    // set user data
    var userName = document.getElementById('userName');
    userName.innerText = user.displayName;
    var userEmail = document.getElementById('userEmail');
    userEmail.innerText = user.mail;

    // hide login button and show user info
    var signInButton = document.getElementById('signin');
    signInButton.style = "display: none";
    var content = document.getElementById('content');
    content.style = "display: block";
}

async function displayEmail() {
    var emails = await getEmails();
    if (!emails || emails.value.length < 1) {
        return;
    }

    var emailsUl = document.getElementById('emails');
    // clear existing elements
    emailsUl.innerHTML = '';
    emails.value.forEach(email => {
        var emailLi = document.createElement('li');
        emailLi.innerText = `${email.subject} (${new Date(email.receivedDateTime).toLocaleString()})`;
        emailsUl.appendChild(emailLi);
    });
}