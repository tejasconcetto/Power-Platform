Clear(colUsers);
ForAll(
    Office365Groups.ListGroupMembers(galTeams.Selected.id).value As A,
    Collect(
        colUsers,
        {
            Name: A.displayName,
            Email: A.mail,
            Photo: Office365Users.UserPhotoV2(A.id),
            Data: With(
                {
                    wJSON: With(
                        {
                            wJSONI: JSON(
                                Office365Users.UserPhotoV2(A.id),
                                JSONFormat.IncludeBinaryData
                            )
                        },
                        Mid(
                            wJSONI,
                            Find(
                                ",",
                                wJSONI
                            ) + 1,
                            Len(wJSONI) - Find(
                                ",",
                                wJSONI
                            ) - 1
                        )
                    )
                },
                wJSON
            )
        }
    );
    
);
Set(
    gblText,
    Concatenate(
        ForAll(
            colUsers As A,
            $"<tr><td class='bg-red'>{A.Name}</td><td style='background-color: #0f0'>{A.Email}</td><td><img width='78px' src='data:image/jpeg;base64,{A.Data}' /></td></tr>"
        ).Value
    );
    
);
Set(
    gblHtml,
    "
  <html>
  <head>
    <title>Test Page</title>
    <style>
     .bg-red { background-color: #ff0000; }
     .bg-green { background: #0f0; }
     .bg-blue { background: #00f; }
     table, td, th {
       border-collapse: collapse;
       border: 1px solid #aaa;
       padding: 4px;
     }
     td { background-color: yellow; }
    </style>
  </head>
  <body>
    <table style='border-collapse: collapse;border: 1px solid #aaa; padding: 5px'>
      " & Concat(
        gblText,
        Value
    ) & "
    </table>
  </body>
 </html>
 "
);