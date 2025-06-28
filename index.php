<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PDF Generator</title>
  <style>
    :root {
      --primary: #4f46e5; /* rich indigo */
      --secondary: #ec4899; /* hot pink */
      --accent: #22c55e; /* green */
      --danger: #ef4444; /* red */
      --bg: #f9fafb;
      --text: #1f2937; /* slate */
    }
    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }
    body {
      font-family: 'Inter', sans-serif;
      background: linear-gradient(135deg, #f0f4ff, #fef9ff);
      color: var(--text);
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      justify-content: center;
      align-items: center;
      padding: 40px 20px;
    }
    .container {
      background: white;
      border-radius: 16px;
      box-shadow: 0 16px 40px rgba(0, 0, 0, 0.08);
      padding: 40px;
      width: 100%;
      max-width: 640px;
      animation: fadeIn 0.6s ease-in-out;
    }
    h1 {
      font-size: 2.2rem;
      text-align: center;
      color: var(--primary);
      margin-bottom: 30px;
      font-weight: 700;
    }
    form label {
      display: block;
      margin-top: 20px;
      font-weight: 600;
      color: #374151;
    }
    form input {
      width: 100%;
      padding: 14px 18px;
      border: 2px solid #d1d5db;
      border-radius: 12px;
      margin-top: 8px;
      font-size: 16px;
      transition: all 0.3s ease;
    }
    form input:focus {
      border-color: var(--primary);
      outline: none;
      box-shadow: 0 0 0 4px rgba(79, 70, 229, 0.2);
    }
    form button {
      margin-top: 30px;
      width: 100%;
      padding: 16px;
      font-size: 16px;
      background: linear-gradient(45deg, var(--danger), #f87171);
      border: none;
      color: white;
      border-radius: 14px;
      font-weight: 700;
      cursor: pointer;
      transition: transform 0.2s ease, box-shadow 0.3s ease;
    }
    form button:hover {
      transform: translateY(-2px);
      box-shadow: 0 10px 24px rgba(0, 0, 0, 0.15);
    }
    iframe {
      width: 100%;
      height: 480px;
      border: none;
      margin-top: 30px;
      border-radius: 10px;
      box-shadow: 0 8px 20px rgba(0, 0, 0, 0.08);
    }
    .footer {
      margin-top: 50px;
      text-align: center;
      font-size: 15px;
      color: #6b7280;
    }
    .footer p:first-child {
      font-weight: 700;
      color: var(--primary);
    }
    @keyframes fadeIn {
      from { opacity: 0; transform: translateY(20px); }
      to { opacity: 1; transform: translateY(0); }
    }
    @media (max-width: 600px) {
      .container {
        padding: 24px;
      }
      h1 {
        font-size: 1.6rem;
      }
    }
  </style>
  <script>
    function validateForm() {
      const inputs = document.querySelectorAll('form input');
      for (let input of inputs) {
        if (!input.value.trim()) {
          alert(`Please fill the '${input.name}' field.`);
          input.focus();
          return false;
        }
      }
      return true;
    }
  </script>
</head>
<body>

  <div class="container">
    <h1>📄 Create Your Assignment PDF</h1>

    <?php
    ini_set('display_errors', 0);
    error_reporting(0);

    $python = "/home/ubuntu/venv/bin/python3";
    $script = __DIR__ . "/fill.py";
    $cmd = "$python $script --detect";
    $output = [];
    exec("$cmd 2>&1", $output, $status);
    $placeholders = array_filter(array_map('trim', $output));

    if ($_SERVER['REQUEST_METHOD'] == 'POST') {
        $args = '';
        foreach ($placeholders as $key) {
            if (isset($_POST[$key]) && trim($_POST[$key]) !== '') {
                $value = escapeshellarg($_POST[$key]);
                $args .= " --$key=$value";
            }
        }
        $cmd = "$python $script $args";
        shell_exec($cmd);

        $pdf_file = __DIR__ . '/zxcvbnm/output_filled.pdf';
        if (file_exists($pdf_file)) {
            echo "<p style='text-align:center; margin-top:20px; font-weight:bold;'><a href='zxcvbnm/output_filled.pdf' download style='color: var(--accent); font-size: 18px;'>⬇️ Download PDF</a></p>";
            echo "<iframe src='zxcvbnm/output_filled.pdf'></iframe>";
        } else {
            echo "<p style='color:red; text-align:center;'>PDF generation failed.</p>";
        }
        exit;
    }

    if (empty($placeholders)) {
        echo "<p style='color:red; text-align:center;'>No placeholders detected in DOCX file.</p>";
    } else {
        echo "<form method='POST' onsubmit='return validateForm()'>";
        foreach ($placeholders as $ph) {
            echo "<label for='$ph'>$ph:</label><input name='$ph' id='$ph'>";
        }
        echo "<button type='submit'>🚀 Generate PDF</button></form>";
    }
    ?>
  </div>

  <div class="footer">
    <p>CyberDeveloper Tech</p>
    <p>bsilent@developer</p>
  </div>

</body>
</html>

