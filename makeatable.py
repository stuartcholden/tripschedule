import csv, smtplib, ssl, xlrd, datetime, subprocess

from_address = "kandalore.trippers@gmail.com"
password = "shalominthehome"
to_address = "tripdirector@kandalore.com"

table = open("tripschedulehtmltable.html","w+")

with open("email.csv",encoding="utf-8") as file:
    reader = csv.DictReader(file)
    for row in reader:
        if row['TripID'] != "":
            staff1 = row['Tripper 1'] + "\n"
        else:
            staff1 = "."

            coremessage = staff1

        table.write("<table class="table table-striped">
  <thead>
    <tr>
      <th scope="col">#</th>
      <th scope="col">First</th>
      <th scope="col">Last</th>
      <th scope="col">Handle</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th scope="row">1</th>
      <td>Mark</td>
      <td>Otto</td>
      <td>@mdo</td>
    </tr>
    <tr>
      <th scope="row">2</th>
      <td>Jacob</td>
      <td>Thornton</td>
      <td>@fat</td>
    </tr>
    <tr>
      <th scope="row">3</th>
      <td>Larry</td>
      <td>the Bird</td>
      <td>@twitter</td>
    </tr>
  </tbody>
</table>")
