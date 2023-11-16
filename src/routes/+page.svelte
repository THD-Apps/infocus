<script>
	import ExcelJS from 'exceljs';
	import moment from 'moment';

	let finishedData = null;
	let loading = false;

	const today = new Date();
	const month = String(today.getMonth() + 1).padStart(2, '0'); // Months are zero-based
	const day = String(today.getDate()).padStart(2, '0');
	const year = String(today.getFullYear()).substr(-2); // Take the last two digits of the year

	const currDate = `${month}/${day}/${year}`;

	function startLoad(event) {
		loading = true;
		setTimeout(() => {
			processFile(event);
		}, 1000);
	}

	function displayFormattedName(fullName) {
		// Split the full name into an array of words
		const nameParts = fullName.split(' ');

		// Extract the first name
		const firstName = nameParts[0];

		// Extract the last name and get the first letter
		const lastName = nameParts.length > 1 ? nameParts[nameParts.length - 1][0] : '';

		// Concatenate the first name and the first letter of the last initial
		const formattedName = `${firstName} ${lastName}.`;

		return formattedName;
	}

	function processFile(event) {
		const file = event.target.files[0];

		if (file) {
			const reader = new FileReader();

			reader.onload = async function (event) {
				const arrayBuffer = event.target.result;

				// Use exceljs to read the Excel file
				const workbook = new ExcelJS.Workbook();
				await workbook.xlsx.load(arrayBuffer);

				const worksheet = workbook.worksheets[0];
				let jsonData = [];

				worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
					const rowJson = {};
					row.eachCell((cell, colNumber) => {
						rowJson[colNumber] = cell.value;
					});
					jsonData.push(rowJson);
				});
				jsonData.sort((a, b) => {
					// Compare by property "2" first
					const property2A = isNaN(a['2']) ? a['2'] : parseInt(a['2']);
					const property2B = isNaN(b['2']) ? b['2'] : parseInt(b['2']);

					if (property2A < property2B) return -1;
					if (property2A > property2B) return 1;

					// If property "2" is the same, then compare by property "1" as strings
					const property1A = a['1'].toUpperCase(); // Convert to uppercase for case-insensitive sorting
					const property1B = b['1'].toUpperCase();

					if (property1A < property1B) return -1;
					if (property1A > property1B) return 1;

					return 0;
				});
				finishedData = jsonData;
				loading = false;
			};

			reader.readAsArrayBuffer(file);
		}
	}
</script>

<div class="text-center">
	{#if loading}
		<div class="font-medium mt-5 text-[25px]">
			Hang tight! Our data hamsters are running as fast as they can...
		</div>
		<div class="hamster text-[200px]">🐹</div>
	{:else if finishedData}
		<div class="flex justify-center screen-only">
			<div class="w-2/5 grid grid-cols-2 gap-3">
				<div
					class="cursor-pointer text-white my-3 bg-blue-700 hover:bg-blue-800 focus:outline-none focus:ring-4 focus:ring-blue-300 font-medium rounded-full text-sm w-50 py-2.5 text-center mb-2 dark:bg-blue-600 dark:hover:bg-blue-700 dark:focus:ring-blue-800"
					on:click={() => window.print()}
				>
					Print Report
				</div>
				<div
					class="cursor-pointer text-white my-3 bg-blue-700 hover:bg-blue-800 focus:outline-none focus:ring-4 focus:ring-blue-300 font-medium rounded-full text-sm w-50 py-2.5 text-center mb-2 dark:bg-blue-600 dark:hover:bg-blue-700 dark:focus:ring-blue-800"
					on:click={() => window.location.reload()}
				>
					Start Over
				</div>
			</div>
		</div>
		<div class="header print-only mx-[30%] print:mx-0">
			<div class="text-[30px]">IN FOCUS NON-PARTICIPANTS</div>

			<div class="text-[22px] font-bold my-3">
				The following associates have not completed the In-Focus Quiz as of {moment(
					currDate
				).format('dddd, MMMM Do')}. Please complete your In-Focus quiz at your earliest possible
				convenience!
			</div>
			<div class="bg-red-400 text-red-800 p-3 text-[20px] mb-4">
				REMEMBER: In-Focus Quizzes should be completed as close to the beginning of each month as
				possible and <b>NO LATER</b> than the 10th of each month.
			</div>
		</div>
		<div class="text-[22px] print:text-[18px] mx-[30%] print:mx-0" id="display-area">
			<div class="text-left">
				{#each finishedData as participant}
					{#if participant['2'] != 'Department'}
						<div>
							D{(participant['2'].length < 2 ? '0' : '') +
								participant['2'] +
								' ' +
								displayFormattedName(participant['1'])}
						</div>
					{/if}
				{/each}
			</div>
		</div>
	{:else}
		<div class="mb-2 mt-4 flex justify-center">
			<img
				width="90"
				height="90"
				src="https://corporate.homedepot.com/sites/default/files/image_gallery/THD_logo.jpg"
				alt=""
			/>
		</div>
		<div class="text-[20pt] mb-3 font-bold text-orange-500">
			In-Focus Non-Participants Report Generator
		</div>
		<div class="flex justify-center">
			<div class="w-2/5 bg-orange-200 py-3">
				<input type="file" id="excel-file" accept=".xlsx, .xls" on:change={startLoad} />
			</div>
		</div>

		<div class="text-xl text-orange-400 mt-3">
			Upload .xlsx In-Focus Report above to continue...
		</div>
	{/if}
</div>

<style>
	.hamster {
		display: inline-block;
		animation: bounce 1s ease-in-out infinite;
	}

	@keyframes bounce {
		0%,
		100% {
			transform: translateY(0);
		}
		50% {
			transform: translateY(-20px);
		}
	}
	@media print {
		.page-break {
			page-break-after: always;
		}
	}
	@media screen {
		.print-only {
			display: none;
		}
	}
	@media print {
		.screen-only {
			display: none;
		}
	}
	@media print {
		#display-area {
			columns: 3; /* Set the number of columns for printing */
		}
		.column-break {
			break-inside: avoid-column; /* Avoid widows by starting a new group on a new column */
		}
	}
</style>
