use fastcdc::v2020::{Normalization, StreamCDC};
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::env;
use std::fs::{self, File};
use std::path::PathBuf;
use std::process;
use std::time::Instant;

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ChunkRecord {
    offset: u64,
    sha256: String,
    bytes: usize,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ChunkConfiguration {
    minimum_bytes: usize,
    average_bytes: usize,
    maximum_bytes: usize,
    elapsed_ms: u128,
    chunks: Vec<ChunkRecord>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct FileIndex {
    path: String,
    bytes: u64,
    configurations: Vec<ChunkConfiguration>,
}

fn fail(message: &str) -> ! {
    eprintln!("{message}");
    process::exit(1);
}

fn digest_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(digest.len() * 2);
    for byte in digest {
        use std::fmt::Write;
        write!(&mut output, "{byte:02x}").expect("writing to String cannot fail");
    }
    output
}

fn parse_arguments() -> (PathBuf, Vec<usize>) {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .map(PathBuf::from)
        .unwrap_or_else(|| fail("usage: onlyoffice-fastcdc-index <pack> [average-bytes ...]"));
    let mut averages = Vec::new();
    for value in arguments {
        let average = value
            .parse::<usize>()
            .unwrap_or_else(|_| fail(&format!("invalid average chunk size: {value}")));
        if !(256..=4_194_304).contains(&average) {
            fail("average chunk size must be between 256 and 4194304 bytes");
        }
        averages.push(average);
    }
    if averages.is_empty() {
        averages.extend([262_144, 1_048_576, 4_194_304]);
    }
    (input, averages)
}

fn main() {
    let (input, averages) = parse_arguments();
    let metadata = fs::metadata(&input)
        .unwrap_or_else(|error| fail(&format!("failed to stat {}: {error}", input.display())));
    let mut configurations = Vec::new();

    for average_bytes in averages {
        let minimum_bytes = average_bytes / 4;
        let maximum_bytes = average_bytes * 4;
        let source = File::open(&input)
            .unwrap_or_else(|error| fail(&format!("failed to open {}: {error}", input.display())));
        let started_at = Instant::now();
        let mut chunks = Vec::new();
        for result in StreamCDC::with_level_and_seed(
            source,
            minimum_bytes,
            average_bytes,
            maximum_bytes,
            Normalization::Level1,
            0,
        ) {
            let chunk = result.unwrap_or_else(|error| {
                fail(&format!("failed to index {}: {error}", input.display()))
            });
            chunks.push(ChunkRecord {
                offset: chunk.offset,
                sha256: digest_hex(&chunk.data),
                bytes: chunk.length,
            });
        }
        configurations.push(ChunkConfiguration {
            minimum_bytes,
            average_bytes,
            maximum_bytes,
            elapsed_ms: started_at.elapsed().as_millis(),
            chunks,
        });
    }

    let output = FileIndex {
        path: input.display().to_string(),
        bytes: metadata.len(),
        configurations,
    };
    println!(
        "{}",
        serde_json::to_string(&output).expect("serializing the FastCDC index cannot fail")
    );
}
