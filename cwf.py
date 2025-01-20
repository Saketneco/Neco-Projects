from flask import Flask, render_template, request, redirect, url_for, flash, session, g ,jsonify
from flask_mysqldb import MySQL
import os
import binascii
from werkzeug.security import check_password_hash

app = Flask(__name__)

# Configure the MySQL connection (for WAMP)
app.config['MYSQL_HOST'] = 'localhost'  # WAMP default MySQL host
app.config['MYSQL_USER'] = 'root'  # Default MySQL user
app.config['MYSQL_PASSWORD'] = ''  # Default password (empty string in WAMP)
app.config['MYSQL_DB'] = 'jnayil_db'  # Replace with your actual database name
app.config['SECRET_KEY'] = os.urandom(24)  # For CSRF protection
mysql = MySQL(app)


@app.before_request
def before_request():
    if 'customer_id' in session:
        g.user = session['customer_id']  # Set g.user to the logged-in user's customer_id
    else:
        g.user = None  # No user logged in


# Route for the login page
@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        # CSRF Token check
        csrf_token = request.form['csrf_token']
        if csrf_token != session.get('csrf_token'):
            flash("Invalid CSRF token. Please try again.", "danger")
       
        
        customer_id = request.form['customer_id']
        password = request.form['password']
        
        g.user = customer_id
        
        # Query the database for user authentication
        cur = mysql.connection.cursor()
        cur.execute("SELECT password FROM users WHERE customer_id = %s", (customer_id,))
        user = cur.fetchone()
        print(f"user:{user}")
        
        if user[0] == password :
            # Store user session
            session['customer_id'] = customer_id
            flash("Login successful!", "success")
            return redirect(url_for('dashboard'))
        else:
            flash("Invalid credentials. Please try again.", "danger")
    
    # Generate a CSRF token and store it in the session
    session['csrf_token'] = binascii.hexlify(os.urandom(32)).decode()
    
    return render_template('login.html')

# Route for the dashboard after login
@app.route('/dashboard')
def dashboard():
    if 'customer_id' not in session:
        return redirect(url_for('login'))
    
    return render_template('landing_page.html')

@app.route('/logout')
def logout():
    return render_template('login.html')

from flask import Response
import json

@app.route('/fetch_invoices')
def fetch_invoices():
    cur = mysql.connection.cursor()
    # Execute the SQL query and fetch all rows
    cur.execute("SELECT customer_id, invoice_number, status FROM pending_invoices WHERE customer_id = %s", (g.user,))
    invoices = cur.fetchall()  # fetchall retrieves all rows

    # Process the data to format it into a list of dictionaries
    invoice_data = [
        {"invoice_number": invoice[1], "customer_id": invoice[0], "status": invoice[2]}
        for invoice in invoices
    ]
    
    # Close the cursor after fetching the data
    cur.close()

    # Manually serialize the data to JSON and return as a Response object
    response_data = json.dumps(invoice_data)

    return Response(response_data, mimetype='application/json')  # Return as JSON with the correct MIME type

from flask import Response
import json
import MySQLdb  # or use your database connector as per your setup

@app.route('/fetch_invoice_details/<invoice_number>', methods=['GET'])
def fetch_invoice_details(invoice_number):
    cur = mysql.connection.cursor()
    
    # SQL query to fetch the invoice details by invoice_number
    query = """
        SELECT 

            invoice_number, 
            invoice_item,
           
            customer_id, 
            customer_name, 
            invoice_quantity, 
            received_quantity,
            material,
            material_description,
            sales_unit,
            remark
        FROM invoices
        WHERE invoice_number = %s
    """
    
    # Execute the query with the invoice number as a parameter
    cur.execute(query, (invoice_number,))
    
    # Fetch all the results (this will be a list of rows)
    invoices = cur.fetchall()
    print(invoices)
    
    # Close the cursor
    cur.close()

    # If no invoices are found, return a 404 error response
    if not invoices:
        response_data = json.dumps({"error": "Invoices not found"})
        return Response(response_data, status=404, mimetype='application/json')
    
    # Prepare a list of dictionaries to hold all invoice details
    invoice_data_list = []
    for invoice in invoices:
        invoice_data = {

            'invoice_number':invoice[0], 
            'invoice_item':invoice[1],
            
            'customer_id':invoice[2], 
            'customer_name':invoice[3], 
            'invoice_quantity':invoice[4], 
            'received_quantity':invoice[5],
            'material':invoice[6],
            'material_description': invoice[7],
            'sales_unit': invoice[8],
            'remark':invoice[9]
            
        }
        invoice_data_list.append(invoice_data)
    
    # Convert the list of invoices to a JSON string
    response_data = json.dumps(invoice_data_list)
    
    # Return the invoice details as JSON response
    return Response(response_data, status=200, mimetype='application/json')


@app.route('/update_invoice_details/<invoice_number>', methods=['POST'])
def update_invoice_details(invoice_number):
    try:
        # Get the JSON data from the request
        data = request.get_json()
        
        
        print(data)
        # Extract the invoice details
        invoice_number = data.get('invoice_number')
        details = data.get('details')

        if not invoice_number or not details:
            return jsonify({"status": "error", "message": "Invoice number or details missing"}), 400

        # Create a cursor to execute SQL queries
        cur = mysql.connection.cursor()

        # Iterate over each detail and update the corresponding invoice details
        for detail in details:
            received_quantity = detail.get('received_quantity')
            remark = detail.get('remark')
            

            # SQL query to update the invoice details
            update_query = """
                UPDATE invoices
                SET received_quantity = %s, remark = %s
                WHERE invoice_number = %s 
            """

            # Execute the update query
            cur.execute(update_query, (received_quantity, remark, invoice_number))

        # Commit the changes to the database
        mysql.connection.commit()

        # Close the cursor
        cur.close()

        # Respond with a success message
        return jsonify({"status": "success", "message": "Invoice details updated successfully."})
    
    except MySQLdb.MySQLError as e:
        # If something goes wrong with the database
        mysql.connection.rollback()
        return jsonify({"status": "error", "message": str(e)}), 400
    
    except Exception as e:
        # General error handling
        return jsonify({"status": "error", "message": str(e)}), 400
        

# Main block to run the Flask application
if __name__ == '__main__':
    app.run(host='0.0.0.0' ,debug=True)
