import glob
import numpy as np
import plotly.graph_objects as go
import skrf as rf
import os
import argparse

def param_to_str(key, value):
    """
    Converts a parameter key-value pair into a string format for labeling.
    by default, it formats as "key=value". 
    For key "Box" (case insensitive), it formats as "Box#value".
    For key "Delta" (case insensitive), it formats with the Delta sign as "Δ=value".
    """
    if key.lower() == "box":
        return f"Box#{value}"
    elif key.lower() == "delta":
        return f"Δ={value}"
    return f"{key}={value}"


def analyze_s2p_files(file_pattern, show_legend=True, save_path=None):
    """
    Visualizes S21 for all .s2p files matching the wildcard pattern
    and identifies the frequency of the resonance dip.
    """
    # Find all files matching the wildcard pattern
    files = glob.glob(file_pattern)
    
    if not files:
        print(f"No files found matching pattern: {file_pattern}")
        return

    fig = go.Figure()

    for file_path in files:
        # Load the Touchstone file
        # skrf.Network automatically parses frequencies and S-parameters
        ntwk = rf.Network(file_path)
        filename = os.path.basename(file_path)

        # S21 is the transmission from Port 1 to Port 2
        # We use .s_db to get the magnitude in decibels
        freqs = ntwk.f  # Frequencies in Hz
        s21_db = ntwk.s_db[:, 1, 0] # Index [1,0] corresponds to S21

        # Find the frequency of the dip (minimum S21 value)
        dip_idx = np.argmin(s21_db)
        dip_freq = freqs[dip_idx]
        dip_val = s21_db[dip_idx]

        print(f"File: {filename:25} | Dip at: {dip_freq/1e9:8.4f} GHz | Magnitude: {dip_val:6.2f} dB")

        # Add S21 magnitude trace
        parameters = filename.split('_')[0] # Extract parameter from filename
        # check if the parameters are in the format "param1=value1,param2=value2"
        if ',' in parameters:
            param_dict = dict(param.split('=') for param in parameters.split(','))
            param_str = ', '.join(param_to_str(k, v) for k, v in param_dict.items())
        else:
            param_str = parameters

        fig.add_trace(go.Scatter(
            x=freqs / 1e9,
            y=s21_db,
            mode='markers',
            marker_size=4,
            name=f"{param_str:<20}\t(Dip: {dip_freq/1e9:.3f} GHz)"
        ))

        # Add a marker at the dip location
        fig.add_trace(go.Scatter(
            x=[dip_freq / 1e9],
            y=[dip_val],
            mode='markers',
            marker=dict(size=10, color='red', symbol='x'),
            showlegend=False,
            hoverinfo='text',
            text=f"Dip: {dip_freq/1e9:.4f} GHz<br>Value: {dip_val:.2f} dB"
        ))

    fig.update_layout(
        title="VNA Measurement: S₂₁ Magnitude",
        xaxis_title="Frequency (GHz)",
        yaxis_title="Magnitude (dB)",
        template="plotly_white",
        hovermode="closest",
        showlegend=show_legend
    )

    if save_path:
        if save_path.endswith('.html'):
            fig.write_html(save_path)
        else:
            fig.write_image(save_path)
        print(f"Plot saved to: {save_path}")

    fig.show()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Visualize S21 for all .s2p files matching a pattern.")
    parser.add_argument("pattern", help="Wildcard pattern for .s2p files (e.g., 'data/*.s2p')")
    parser.add_argument("--hide-legend", action="store_true", help="Hide the legend in the plot")
    parser.add_argument("--save-as", type=str, help="Save the figure to a file (e.g., 'output.html' or 'output.png')")

    args = parser.parse_args()
    analyze_s2p_files(args.pattern, show_legend=not args.hide_legend, save_path=args.save_as)
