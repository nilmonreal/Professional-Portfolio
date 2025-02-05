from flask import Flask, render_template, request, redirect, url_for, flash
from forms import JobOfferForm
from flask_mail import Mail, Message
from datetime import datetime
from flask_wtf import FlaskForm
from wtforms import StringField, TextAreaField, SubmitField
from wtforms.validators import DataRequired, Email, Length

class JobOfferForm(FlaskForm):
    name = StringField('Name', validators=[DataRequired(), Length(max=50)])
    email = StringField('Email', validators=[DataRequired(), Email(), Length(max=120)])
    message = TextAreaField('Message', validators=[DataRequired(), Length(max=500)])
    submit = SubmitField('Send')

app = Flask(__name__)
app.secret_key = 'nifu'

# Optional: Configure Flask-Mail if you want to send emails directly from the form
app.config['MAIL_SERVER'] = 'smtp.gmail.com'  # Replace with your mail server
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USERNAME'] = 'nilmonreal@gmail.com'
app.config['MAIL_PASSWORD'] = 'rtld ohli nzgd mzsr'
mail = Mail(app)


projects = [
    {
        'title': 'Morse to Code Converter',
        'description': 'Code to decrypt morse lenguage.',
        'github_url': 'https://github.com/yourusername/ecommerce-platform',
        'technologies': ['Python'],
        'images': ['image1.jpg', 'morse-codes.jpg']
    },
    {
        'title': 'Data Visualization Dashboard',
        'description': 'An interactive dashboard to visualize real-time data using Flask and D3.js.',
        'github_url': 'https://github.com/yourusername/data-visualization-dashboard',
        'technologies': ['Flask', 'D3.js'],
        'images': ['image1.jpg', 'image2.jpg']
    },
    {
        'title': 'Mobile To-Do App',
        'description': 'A cross-platform mobile application to manage tasks efficiently, built using Flutter.',
        'github_url': 'https://github.com/yourusername/mobile-todo-app',
        'technologies': ['Flutter'],
        'images': ['image1.jpg', 'image2.jpg']
    }
]

@app.route('/')
def home():
    return render_template('index.html', title='Home', projects=projects, current_year = datetime.now().year)

@app.route('/projects')
def projects_page():
    return render_template('projects.html', title='Projects', projects=projects, current_year = datetime.now().year)

@app.route('/projects/<string:title>')
def project_detail(title):
    project = next((proj for proj in projects if proj['title'] == title), None)
    if project is None:
        flash('Project not found.', 'danger')
        return redirect(url_for('projects_page'))
    return render_template('project_detail.html', title=project['title'], project=project, current_year = datetime.now().year)


# @app.route('/about', methods=['GET', 'POST'])
# def about():
#     form = JobOfferForm()
#     if form.validate_on_submit():
#         # Handle form submission logic, e.g., send an email or store in a database
#         name = form.name.data
#         email = form.email.data
#         message = form.message.data

#         # Send an email (optional)
#         msg = Message('Job Offer from Portfolio', sender=email, recipients=['nilmonreal@gmail.com'])
#         msg.body = f"Name: {name}\nEmail: {email}\n\n{message}"
#         mail.send(msg)

#         # Optional: Send confirmation email to the user
#         confirmation_msg = Message('Thank you for your message!', sender='your_email@example.com', recipients=[email])
#         confirmation_msg.body = f"Hi {name},\n\nThank you for reaching out with a job offer. I'll get back to you soon.\n\nBest regards,\nYour Name"
#         mail.send(confirmation_msg)

#         flash('Thank you! Your message has been sent, and we have emailed a confirmation.', 'success')
#         return redirect(url_for('about'))
#     return render_template('about.html', title='About Me', form=form, current_year = datetime.now().year)

@app.route('/about', methods=['GET', 'POST'])
def about():
    form = JobOfferForm()
    if form.validate_on_submit():
        # Send email
        msg = Message(
            subject="Job Offer from Portfolio",
            sender=form.email.data,
            recipients=['nilmonreal@gmail.com']
        )
        msg.body = f"Name: {form.name.data}\nEmail: {form.email.data}\n\nMessage:\n{form.message.data}"
        mail.send(msg)

        flash('Thank you! Your message has been sent.', 'success')
        return redirect(url_for('about'))
    return render_template('about.html', title='About Me', form=form, current_year = datetime.now().year)



if __name__ == '__main__':
    app.run(debug=True, use_reloader=False)
