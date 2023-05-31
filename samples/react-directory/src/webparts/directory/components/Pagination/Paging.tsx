import * as React from 'react';
import { useEffect, useState } from 'react';
import Pagination from 'react-js-pagination';
import styles from './Paging.module.scss';

export type PageUpdateCallback = (pageNumber: number) => void;

export interface IPagingProps {
	totalItems: number;
	itemsCountPerPage: number;
	onPageUpdate: PageUpdateCallback;
	currentPage: number;
}

export interface IPagingState {
	currentPage: number;
}

const Paging: React.FC<IPagingProps> = (props) => {
	const [currentPage, setcurrentPage] = useState<number>(props.currentPage);

	const _pageChange = (pageNumber: number): void => {
		setcurrentPage(pageNumber);
		props.onPageUpdate(pageNumber);
	};

	useEffect(() => {
		setcurrentPage(props.currentPage);
	}, [props.currentPage]);

	return (
		<div className={styles.paginationContainer}>
			<div className={styles.searchWp__paginationContainer__pagination}>
				<Pagination
					activePage={currentPage}
					// eslint-disable-next-line react/self-closing-comp
					firstPageText={<i className='ms-Icon ms-Icon--DoubleChevronLeft' aria-hidden='true'></i>}
					// eslint-disable-next-line react/self-closing-comp
					lastPageText={<i className='ms-Icon ms-Icon--DoubleChevronRight' aria-hidden='true'></i>}
					// eslint-disable-next-line react/self-closing-comp
					prevPageText={<i className='ms-Icon ms-Icon--ChevronLeft' aria-hidden='true'></i>}
					// eslint-disable-next-line react/self-closing-comp
					nextPageText={<i className='ms-Icon ms-Icon--ChevronRight' aria-hidden='true'></i>}
					activeLinkClass={styles.active}
					itemsCountPerPage={props.itemsCountPerPage}
					totalItemsCount={props.totalItems}
					pageRangeDisplayed={5}
					onChange={_pageChange}
				/>
			</div>
		</div>
	);
};

export default Paging;
