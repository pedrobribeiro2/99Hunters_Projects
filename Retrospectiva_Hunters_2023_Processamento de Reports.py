{
  "nbformat": 4,
  "nbformat_minor": 0,
  "metadata": {
    "colab": {
      "provenance": [],
      "authorship_tag": "ABX9TyNeiGa2XgxRfQysKhZp6kH/"
    },
    "kernelspec": {
      "name": "python3",
      "display_name": "Python 3"
    },
    "language_info": {
      "name": "python"
    }
  },
  "cells": [
    {
      "cell_type": "code",
      "source": [
        "from google.colab import drive\n",
        "drive.mount('/content/drive')\n"
      ],
      "metadata": {
        "id": "y6oThW7ohShb",
        "colab": {
          "base_uri": "https://localhost:8080/"
        },
        "outputId": "05d08e74-3aa9-49b0-d510-d4b927a5d024"
      },
      "execution_count": null,
      "outputs": [
        {
          "output_type": "stream",
          "name": "stdout",
          "text": [
            "Drive already mounted at /content/drive; to attempt to forcibly remount, call drive.mount(\"/content/drive\", force_remount=True).\n"
          ]
        }
      ]
    },
    {
      "cell_type": "code",
      "execution_count": null,
      "metadata": {
        "id": "9zawrb_MfYN-",
        "colab": {
          "base_uri": "https://localhost:8080/"
        },
        "outputId": "6c1869a5-a940-4bfb-a1ae-62e20acbb4d1"
      },
      "outputs": [
        {
          "output_type": "stream",
          "name": "stdout",
          "text": [
            "   ID_Hunter_Job  ID_Hunter  ID_Job       Job Executive      Job Manager  \\\n",
            "0          26405       2719    4477  Danielle Calheiros  Victor Valentim   \n",
            "1          26406       4294    4477  Danielle Calheiros  Victor Valentim   \n",
            "2          26407       4442    4477  Danielle Calheiros  Victor Valentim   \n",
            "3          26408       3286    4477  Danielle Calheiros  Victor Valentim   \n",
            "4          26409       6032    4477  Danielle Calheiros  Victor Valentim   \n",
            "\n",
            "                        Hunter_Name Hunter is Team99 Company_Name  \\\n",
            "0                      Pamela Simas           Hunter          GPA   \n",
            "1                     Jessica Silva           Hunter          GPA   \n",
            "2                    Luana Carvalho           Hunter          GPA   \n",
            "3  Chiara Carolina Martins de Souza           Hunter          GPA   \n",
            "4                         Giselle .           Hunter          GPA   \n",
            "\n",
            "                                     Job_Name  Job Code  ... recommended  \\\n",
            "0  Pessoa Coordenadora de Planejamento Sênior    216101  ...         1.0   \n",
            "1  Pessoa Coordenadora de Planejamento Sênior    216101  ...         2.0   \n",
            "2  Pessoa Coordenadora de Planejamento Sênior    216101  ...         0.0   \n",
            "3  Pessoa Coordenadora de Planejamento Sênior    216101  ...         2.0   \n",
            "4  Pessoa Coordenadora de Planejamento Sênior    216101  ...         NaN   \n",
            "\n",
            "  shortlisted  interviewed hired HunterScore Hunter total Points Fee  \\\n",
            "0         0.0          0.0   0.0          31                4634   0   \n",
            "1         0.0          0.0   0.0          26                 108   0   \n",
            "2         0.0          0.0   0.0          12                  87   0   \n",
            "3         0.0          0.0   0.0          24                 885   0   \n",
            "4         NaN          NaN   NaN          37                  50   0   \n",
            "\n",
            "  First Open Job  Time to First Reco Deleted_Hunter  \n",
            "0          False                 4.0          False  \n",
            "1          False                 2.0          False  \n",
            "2          False                 NaN          False  \n",
            "3          False                 7.0          False  \n",
            "4          False                 NaN          False  \n",
            "\n",
            "[5 rows x 31 columns]\n"
          ]
        }
      ],
      "source": [
        "import pandas as pd\n",
        "\n",
        "# Adjust the file path below to the location of your file in Google Drive\n",
        "file_path = '/content/drive/MyDrive/Pedro Marketing/2023 12 99Hunters Wrapped/2023_12_15_hunter_jobs_report.csv'  # Replace with your file path in Google Drive\n",
        "\n",
        "# Load the dataset\n",
        "data = pd.read_csv(file_path)\n",
        "\n",
        "# Display the first few rows to verify the data\n",
        "print(data.head())\n"
      ]
    },
    {
      "cell_type": "code",
      "source": [
        "# Group the data by 'ID_Hunter' and calculate the necessary statistics\n",
        "grouped_data = data.groupby('ID_Hunter').agg(\n",
        "    Full_Name=('Hunter_Name', 'first'),  # Get the first name in each group\n",
        "    Total_Recommendations=('recommended', 'sum'),\n",
        "    Total_Shortlisted=('shortlisted', 'sum'),\n",
        "    Total_Interviewed=('interviewed', 'sum'),\n",
        "    Total_Hired=('hired', 'sum'),\n",
        "    Earnings_2023=('Fee', 'sum')  # Assuming 'Fee' column represents earnings\n",
        ").reset_index()\n",
        "\n",
        "# Display the processed data\n",
        "print(grouped_data.head())\n"
      ],
      "metadata": {
        "id": "LLwW5ZL7g0Fk",
        "colab": {
          "base_uri": "https://localhost:8080/"
        },
        "outputId": "600bbf48-4b5e-4dce-df7c-72f009a8f1ff"
      },
      "execution_count": null,
      "outputs": [
        {
          "output_type": "stream",
          "name": "stdout",
          "text": [
            "   ID_Hunter               Full_Name  Total_Recommendations  \\\n",
            "0          4        Luciano Montezzo                   28.0   \n",
            "1         46           Maib Oliveira                    3.0   \n",
            "2         63           Felipe Arruda                    0.0   \n",
            "3         65          Cristina Cappi                    0.0   \n",
            "4         66  Marta Helena Gonçalves                    0.0   \n",
            "\n",
            "   Total_Shortlisted  Total_Interviewed  Total_Hired  Earnings_2023  \n",
            "0                5.0                2.0          0.0            200  \n",
            "1                1.0                1.0          0.0            100  \n",
            "2                0.0                0.0          0.0              0  \n",
            "3                0.0                0.0          0.0              0  \n",
            "4                0.0                0.0          0.0              0  \n"
          ]
        }
      ]
    },
    {
      "cell_type": "code",
      "source": [
        "# Filtrando os dados para incluir apenas os status desejados\n",
        "status_validos = ['open', 'interviews', 'shortlist', 'placement', 'adv_interviews', 'stand_by']\n",
        "data_filtrada = data[data['HunterJob Status'].isin(status_validos)]\n",
        "\n",
        "# Contabilizando o número de jobs trabalhados por cada Hunter\n",
        "jobs_contabilizados = data_filtrada.groupby('ID_Hunter')['ID_Job'].nunique().reset_index()\n",
        "jobs_contabilizados.rename(columns={'ID_Job': 'Total_Jobs_Trabalhados'}, inplace=True)\n",
        "\n",
        "# Mesclando os dados contabilizados com os dados agrupados anteriores\n",
        "final_df = pd.merge(grouped_data, jobs_contabilizados, on='ID_Hunter', how='left')\n",
        "\n",
        "# Mesclando os dados contabilizados com os dados agrupados anteriores\n",
        "grouped_data = pd.merge(grouped_data, jobs_contabilizados, on='ID_Hunter', how='left')\n",
        "\n",
        "# Preenchendo valores ausentes com 0, caso algum Hunter não tenha jobs contabilizados\n",
        "grouped_data['Total_Jobs_Trabalhados'].fillna(0, inplace=True)\n"
      ],
      "metadata": {
        "id": "EC4-SvRbs0_a"
      },
      "execution_count": null,
      "outputs": []
    },
    {
      "cell_type": "code",
      "source": [
        "# Importar a biblioteca necessária\n",
        "import calendar\n",
        "\n",
        "# Converter a coluna 'Created_at' para datetime\n",
        "data['Created_at'] = pd.to_datetime(data['Created_at'], format='%m/%d/%Y-%H:%M')\n",
        "\n",
        "# Extrair o mês e o ano de cada job\n",
        "data['Month'] = data['Created_at'].dt.month\n",
        "data['Year'] = data['Created_at'].dt.year\n",
        "\n",
        "# Filtrar os dados para incluir apenas os jobs dos meses e status desejados\n",
        "data_filtered = data[(data['HunterJob Status'].isin(status_validos)) & (data['Year'] == 2023)]\n",
        "\n",
        "# Contabilizar os jobs por mês para cada hunter\n",
        "monthly_activity = data_filtered.groupby(['ID_Hunter', 'Month']).size().reset_index(name='Jobs_Per_Month')\n",
        "\n",
        "# Mapeamento de número do mês para nome do mês em português\n",
        "meses_pt_br = {\n",
        "    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril', 5: 'Maio', 6: 'Junho',\n",
        "    7: 'Julho', 8: 'Agosto', 9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'\n",
        "}\n",
        "\n",
        "# Encontrar os três meses mais ativos para cada hunter\n",
        "def top_3_months(df):\n",
        "    top_months = df.sort_values('Jobs_Per_Month', ascending=False).head(3)\n",
        "    # Ordenar os meses cronologicamente antes de juntá-los em uma string\n",
        "    top_months_sorted = top_months.sort_values('Month')\n",
        "    return ', '.join(top_months_sorted['Month'].apply(lambda x: meses_pt_br[x]))\n",
        "\n",
        "# Aplicação da função top_3_months e criação de top_months_df\n",
        "top_months_df = monthly_activity.groupby('ID_Hunter').apply(top_3_months).reset_index(name='Top_3_Active_Months')\n",
        "\n",
        "# Mesclagem com grouped_data\n",
        "grouped_data = pd.merge(grouped_data, top_months_df, on='ID_Hunter', how='left')\n",
        "\n",
        "# Preenchimento de valores ausentes\n",
        "grouped_data['Top_3_Active_Months'].fillna('No active months', inplace=True)"
      ],
      "metadata": {
        "id": "DKUSEcj22zfv"
      },
      "execution_count": null,
      "outputs": []
    },
    {
      "cell_type": "code",
      "source": [
        "# Rank hunters based on Total Recommendations in descending order\n",
        "grouped_data['Ranking'] = grouped_data['Total_Recommendations'].rank(ascending=False, method='min')\n"
      ],
      "metadata": {
        "id": "SU6_U16ikPTl"
      },
      "execution_count": null,
      "outputs": []
    },
    {
      "cell_type": "code",
      "source": [
        "import numpy as np\n",
        "\n",
        "# Calculate the percentile rank\n",
        "grouped_data['Percentile'] = grouped_data['Total_Recommendations'].rank(pct=True, ascending=False)\n",
        "\n",
        "# Convert the percentile to a more readable format (e.g., top 10%)\n",
        "grouped_data['Percentile'] = np.ceil(grouped_data['Percentile'] * 100).astype(int)\n"
      ],
      "metadata": {
        "id": "H9zf3YEwkbgn"
      },
      "execution_count": null,
      "outputs": []
    },
    {
      "cell_type": "code",
      "source": [
        "# Converter colunas para tipo inteiro\n",
        "colunas_numericas = ['Total_Jobs_Trabalhados', 'Total_Recommendations', 'Total_Shortlisted', 'Total_Interviewed', 'Total_Hired', 'Ranking']\n",
        "grouped_data[colunas_numericas] = grouped_data[colunas_numericas].fillna(0).astype(int)\n",
        "\n",
        "# Formatar 'Earnings_2023' como moeda BRL\n",
        "grouped_data['Earnings_2023'] = grouped_data['Earnings_2023'].apply(lambda x: f'R$ {x:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.'))\n",
        "\n",
        "# Verificar os tipos de dados após conversão\n",
        "print(grouped_data.dtypes)\n"
      ],
      "metadata": {
        "id": "0bX-kM46KZ0O",
        "colab": {
          "base_uri": "https://localhost:8080/"
        },
        "outputId": "9ab345fb-8da8-44d7-966d-2f1b93b6fd1f"
      },
      "execution_count": null,
      "outputs": [
        {
          "output_type": "stream",
          "name": "stdout",
          "text": [
            "ID_Hunter                  int64\n",
            "Full_Name                 object\n",
            "Total_Recommendations      int64\n",
            "Total_Shortlisted          int64\n",
            "Total_Interviewed          int64\n",
            "Total_Hired                int64\n",
            "Earnings_2023             object\n",
            "Total_Jobs_Trabalhados     int64\n",
            "Top_3_Active_Months       object\n",
            "Ranking                    int64\n",
            "Percentile                 int64\n",
            "dtype: object\n"
          ]
        }
      ]
    },
    {
      "cell_type": "code",
      "source": [
        "# Create the final DataFrame with the new column\n",
        "final_df = grouped_data[['ID_Hunter', 'Full_Name', 'Total_Jobs_Trabalhados','Total_Recommendations', 'Total_Shortlisted', 'Total_Interviewed', 'Total_Hired', 'Earnings_2023', 'Ranking', 'Percentile', 'Top_3_Active_Months']]\n",
        "\n",
        "# Sort the DataFrame by 'Ranking' in ascending order (top rank first)\n",
        "final_df = final_df.sort_values(by='Ranking', ascending=True)\n",
        "\n",
        "# Export to CSV\n",
        "final_df.to_csv('Retrospectiva_Hunters_2023_12_15_VF.csv', index=False)\n"
      ],
      "metadata": {
        "id": "KOl1e8bBg5-a"
      },
      "execution_count": null,
      "outputs": []
    },
    {
      "cell_type": "code",
      "source": [
        "from google.colab import files\n",
        "files.download('Retrospectiva_Hunters_2023_12_15_VF.csv')\n"
      ],
      "metadata": {
        "id": "XTLpySUAg_au",
        "colab": {
          "base_uri": "https://localhost:8080/",
          "height": 17
        },
        "outputId": "87822d78-86a8-4271-e5fc-8bf255a515f4"
      },
      "execution_count": null,
      "outputs": [
        {
          "output_type": "display_data",
          "data": {
            "text/plain": [
              "<IPython.core.display.Javascript object>"
            ],
            "application/javascript": [
              "\n",
              "    async function download(id, filename, size) {\n",
              "      if (!google.colab.kernel.accessAllowed) {\n",
              "        return;\n",
              "      }\n",
              "      const div = document.createElement('div');\n",
              "      const label = document.createElement('label');\n",
              "      label.textContent = `Downloading \"${filename}\": `;\n",
              "      div.appendChild(label);\n",
              "      const progress = document.createElement('progress');\n",
              "      progress.max = size;\n",
              "      div.appendChild(progress);\n",
              "      document.body.appendChild(div);\n",
              "\n",
              "      const buffers = [];\n",
              "      let downloaded = 0;\n",
              "\n",
              "      const channel = await google.colab.kernel.comms.open(id);\n",
              "      // Send a message to notify the kernel that we're ready.\n",
              "      channel.send({})\n",
              "\n",
              "      for await (const message of channel.messages) {\n",
              "        // Send a message to notify the kernel that we're ready.\n",
              "        channel.send({})\n",
              "        if (message.buffers) {\n",
              "          for (const buffer of message.buffers) {\n",
              "            buffers.push(buffer);\n",
              "            downloaded += buffer.byteLength;\n",
              "            progress.value = downloaded;\n",
              "          }\n",
              "        }\n",
              "      }\n",
              "      const blob = new Blob(buffers, {type: 'application/binary'});\n",
              "      const a = document.createElement('a');\n",
              "      a.href = window.URL.createObjectURL(blob);\n",
              "      a.download = filename;\n",
              "      div.appendChild(a);\n",
              "      a.click();\n",
              "      div.remove();\n",
              "    }\n",
              "  "
            ]
          },
          "metadata": {}
        },
        {
          "output_type": "display_data",
          "data": {
            "text/plain": [
              "<IPython.core.display.Javascript object>"
            ],
            "application/javascript": [
              "download(\"download_22f98b1b-2bc7-4f54-aaea-b10848204c64\", \"Retrospectiva_Hunters_2023_12_15_VF.csv\", 49088)"
            ]
          },
          "metadata": {}
        }
      ]
    }
  ]
}